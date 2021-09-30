from pywintypes import com_error
from pathlib import Path
import time
from imessagefilter import CMessageFilter

from utils import get_acad, get_drawings, Point, get_document, separator, sub_separator, vt_point

weld_joint_diameter = 1.4


def process(path):
    if Path(path).is_dir():
        process_folder(path)
    elif Path(path).is_file():
        process_files([path])


def process_folder(folder):
    print('Processing under %s' % Path(folder).absolute())
    print(separator)
    start_time = time.time()
    files = get_drawings(folder)
    processed = process_files(files)
    time_spent = time.time() - start_time
    print(separator)
    print(f'{processed}/{len(files)} files processed in {time_spent}s.')


def process_files(files):
    processed = 0
    counter = 1
    if files:
        CMessageFilter.register()
        app = get_acad()
        for file in files:
            print(f'{counter}/{len(files)}')
            print(sub_separator)
            processed += process_file(app, file)
            print(sub_separator)
            counter += 1
        CMessageFilter.revoke()

    return processed


def process_file(app, dwg_file: str):
    """
    Simplify iso_drawing for file merging,
    explode proxy objects, and delete hidden objects in target file.
    :param dwg_file:
    :return:
    """
    processed = 0
    print(dwg_file)
    try:
        start_time = time.time()
        doc = get_document(app, dwg_file)
        model_space = doc.ModelSpace
        layers = doc.Layers
        hidden_layers = [layer.Name for layer in layers if not layer.LayerOn]
        i = model_space.Count
        delete_counter = 0
        proxy_counter = 0

        while i > 0:
            i -= 1
            item = model_space.Item(i)
            # find hidden Objects and take them out
            if item.Layer in hidden_layers or not item.Visible:
                item.Delete()
                delete_counter += 1
                continue

            # find Proxy Object and explode (Excluding weld joint, which will lost solid fill after exploding)
            if item.ObjectName == 'AcDbZombieEntity':
                # weld joints detection and replacement
                if not is_weld_joint(model_space, item):
                    # found
                    # proxy object does not support "explode" method, use command instead.
                    doc.SendCommand('explode (handent"%s")\n\n' % item.Handle)
                proxy_counter += 1
                continue

        # take out Proxy Object hidden in dictinaries
        # dictionaries = doc.Dictionaries
        # j = dictionaries.Count
        #
        # while j > 0:
        #     j -= 1
        #     dictionary = dictionaries.Item(j)
        #     if dictionary.ObjectName == 'AcDbZombieObject':
        #         dictionary.Delete()
        #         delete_counter += 1
        #         continue

        print('Deleted %s hidden objects' % delete_counter)
        print('Explode %s proxy objects' % proxy_counter)
        if delete_counter + proxy_counter > 0:
            doc.PurgeAll()
            doc.Save()

        doc.Close(False)
        time_spent = time.time() - start_time
        print('Finished in %ss' % round(time_spent))
        processed = 1
    except com_error as e:
        print('file error (%s), skipped' % e.excepinfo[2])

    return processed


def is_weld_joint(model_space, entity) -> bool:
    # """
    # Return if target entity is weld joint (A solid circle which diameter is about 1.4)
    # Only for iso drawing output by EP3D
    # :param entity:
    # :return:
    # """
    try:
        p1, p2 = entity.GetBoundingBox()
        p1 = Point(*p1)
        p2 = Point(*p2)
        dx = p2.x - p1.x
        dy = p2.y - p1.y
        # print([dx, dy])
        if 1.39 < dx < weld_joint_diameter and 1.39 < dy < weld_joint_diameter:
            center = Point(p1.x + dx / 2, p1.y + dy / 2, 0)
            draw_weld_joint(model_space, center)
            entity.Delete()
            print('Weld joint rebuild.')
            return True
    except com_error as e:
        return False

    return False


def draw_weld_joint(model_space, center: Point):
    offset = 0.02
    radius = weld_joint_diameter / 2
    while radius > offset:
        model_space.AddCircle(vt_point(*center), radius)
        radius -= offset


def main_ui():
    print(
        """
        A batch process tool for purging isogen drawing
        
        Current Folder is:
        %s
        
        Shall we start? (Press N to quit)
        """
        % Path('').absolute()
    )

    user_input = input(':')
    if user_input == 'n' or user_input == 'N':
        exit()
    else:
        process_folder('')


if __name__ == '__main__':
    # process(r'D:\Code\dwg_proxy_purge\test\XDA202100061-0801-PD-11019.dwg')
    main_ui()

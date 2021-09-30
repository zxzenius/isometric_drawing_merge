//
// Refer to
// https://docs.microsoft.com/en-us/previous-versions/ms228772(v=vs.140)
// https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-imessagefilter
// https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/implementing-iunknown-in-c

#include <stdio.h>
#include <objbase.h>
#include <objidl.h>


static IMessageFilterVtbl imp_filter;
static IMessageFilter newFilter = { &imp_filter };
static ULONG m_refcount = 0;
static LPMESSAGEFILTER *lpOldFilter;


static STDMETHODIMP_(ULONG) MessageFilter_AddRef(IMessageFilter * This)
{
  return ++ m_refcount;
}

static STDMETHODIMP_(ULONG) MessageFilter_Release(IMessageFilter * This)
{
  return -- m_refcount;
}

static STDMETHODIMP MessageFilter_QueryInterface(IMessageFilter * This, REFIID riid, LPVOID* lppvObj)
{
  HRESULT hr = S_OK;
  // Validate the object pointer.
  if (!This) {
      hr = ResultFromScode(E_INVALIDARG);
      return hr;
  }
  // Validate other parameters.
  if (!lppvObj) {
      hr = ResultFromScode(E_INVALIDARG);
      return hr;
  }
  // Set the output pointer to NULL.
  *lppvObj = NULL;

  // Check the interface identifier.
  if (memcmp(riid, &IID_IUnknown, sizeof(IID)) &&
      memcmp(riid, &IID_IMessageFilter, sizeof(IID))) {
      hr = ResultFromScode(E_NOINTERFACE);
      return hr;
  }

  // The interface is supported. Increment the reference count and return.
  This->lpVtbl->AddRef(This);
  *lppvObj = This;
  return hr;
}

static STDMETHODIMP_(DWORD) MessageFilter_HandleInComingCall(IMessageFilter * This,
                                                             /* [in] */ DWORD dwCallType,
                                                             /* [in] */ HTASK htaskCaller,
                                                             /* [in] */ DWORD dwTickCount,
                                                             /* [in] */ LPINTERFACEINFO lpInterfaceInfo)
{
  //Return the flag SERVERCALL_ISHANDLED.
  return 0;
}


static STDMETHODIMP_(DWORD) MessageFilter_RetryRejectedCall(IMessageFilter * This,
                                                            /* [in] */ HTASK htaskCallee,
                                                            /* [in] */ DWORD dwTickCount,
                                                            /* [in] */ DWORD dwRejectType)
{
  if (dwRejectType == 2) {
    // flag = SERVERCALL_RETRYLATER.
    // Retry the thread call immediately if return >=0 & <100.
    return 99;
  }
  // Too busy; cancel call.
  return -1;
}

static STDMETHODIMP_(DWORD) MessageFilter_MessagePending(IMessageFilter * This,
                                                         /* [in] */ HTASK htaskCallee,
                                                         /* [in] */ DWORD dwTickCount,
                                                         /* [in] */ DWORD dwPendingType)
{
  //Return the flag PENDINGMSG_WAITDEFPROCESS.
  return 2;
}

// Start the filter.
void register_message_filter()
{
  imp_filter.QueryInterface = MessageFilter_QueryInterface;
  imp_filter.AddRef = MessageFilter_AddRef;
  imp_filter.Release = MessageFilter_Release;
  imp_filter.HandleInComingCall = MessageFilter_HandleInComingCall;
  imp_filter.RetryRejectedCall = MessageFilter_RetryRejectedCall;
  imp_filter.MessagePending = MessageFilter_MessagePending;

  if (CoRegisterMessageFilter(&newFilter, lpOldFilter) != S_OK)
    printf("Register new message filter failed\n");
}

// Done with the filter, close it.
void revoke_message_filter()
{
  CoRegisterMessageFilter(NULL, lpOldFilter);
}

int t_main(int argc, char* argv[])
{
  HRESULT ret;
  ret = CoInitializeEx(0, COINIT_APARTMENTTHREADED);
  register_message_filter();
  revoke_message_filter();
  CoUninitialize();
  return 0;
}

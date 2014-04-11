/* this ALWAYS GENERATED file contains the proxy stub code */


/* File created by MIDL compiler version 5.01.0164 */
/* at Thu Jan 19 17:00:58 2006
 */
/* Compiler settings for opccomn.idl:
    Os (OptLev=s), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcproxy.h> version is high enough to compile this file*/
#ifndef __REDQ_RPCPROXY_H_VERSION__
#define __REQUIRED_RPCPROXY_H_VERSION__ 440
#endif


#include "rpcproxy.h"
#ifndef __RPCPROXY_H_VERSION__
#error this stub requires an updated version of <rpcproxy.h>
#endif // __RPCPROXY_H_VERSION__


#include "opccomn.h"

#define TYPE_FORMAT_STRING_SIZE   113                               
#define PROC_FORMAT_STRING_SIZE   77                                

typedef struct _MIDL_TYPE_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ TYPE_FORMAT_STRING_SIZE ];
    } MIDL_TYPE_FORMAT_STRING;

typedef struct _MIDL_PROC_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ PROC_FORMAT_STRING_SIZE ];
    } MIDL_PROC_FORMAT_STRING;


extern const MIDL_TYPE_FORMAT_STRING __MIDL_TypeFormatString;
extern const MIDL_PROC_FORMAT_STRING __MIDL_ProcFormatString;


/* Object interface: IUnknown, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IOPCShutdown, ver. 0.0,
   GUID={0xF31DFDE1,0x07B6,0x11d2,{0xB2,0xD8,0x00,0x60,0x08,0x3B,0xA1,0xFB}} */


extern const MIDL_STUB_DESC Object_StubDesc;


#pragma code_seg(".orpc")

HRESULT STDMETHODCALLTYPE IOPCShutdown_ShutdownRequest_Proxy( 
    IOPCShutdown __RPC_FAR * This,
    /* [string][in] */ LPCWSTR szReason)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      3);
        
        
        
        if(!szReason)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 12U;
            NdrConformantStringBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                           (unsigned char __RPC_FAR *)szReason,
                                           (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4] );
            
            NdrProxyGetBuffer(This, &_StubMsg);
            NdrConformantStringMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                                         (unsigned char __RPC_FAR *)szReason,
                                         (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4] );
            
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[0] );
            
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCShutdown_ShutdownRequest_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    LPCWSTR szReason;
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    ( LPCWSTR  )szReason = 0;
    RpcTryFinally
        {
        if ( (_pRpcMessage->DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
            NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[0] );
        
        NdrConformantStringUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                       (unsigned char __RPC_FAR * __RPC_FAR *)&szReason,
                                       (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4],
                                       (unsigned char)0 );
        
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCShutdown*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> ShutdownRequest((IOPCShutdown *) ((CStdStubBuffer *)This)->pvServerObject,szReason);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 4U;
        _StubMsg.BufferLength += 16;
        
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}

const CINTERFACE_PROXY_VTABLE(4) _IOPCShutdownProxyVtbl = 
{
    &IID_IOPCShutdown,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    IOPCShutdown_ShutdownRequest_Proxy
};


static const PRPC_STUB_FUNCTION IOPCShutdown_table[] =
{
    IOPCShutdown_ShutdownRequest_Stub
};

const CInterfaceStubVtbl _IOPCShutdownStubVtbl =
{
    &IID_IOPCShutdown,
    0,
    4,
    &IOPCShutdown_table[-3],
    CStdStubBuffer_METHODS
};


/* Object interface: IOPCCommon, ver. 0.0,
   GUID={0xF31DFDE2,0x07B6,0x11d2,{0xB2,0xD8,0x00,0x60,0x08,0x3B,0xA1,0xFB}} */


extern const MIDL_STUB_DESC Object_StubDesc;


#pragma code_seg(".orpc")

HRESULT STDMETHODCALLTYPE IOPCCommon_SetLocaleID_Proxy( 
    IOPCCommon __RPC_FAR * This,
    /* [in] */ LCID dwLcid)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      3);
        
        
        
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 4U;
            NdrProxyGetBuffer(This, &_StubMsg);
            *(( LCID __RPC_FAR * )_StubMsg.Buffer)++ = dwLcid;
            
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[6] );
            
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCCommon_SetLocaleID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    LCID dwLcid;
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    RpcTryFinally
        {
        if ( (_pRpcMessage->DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
            NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[6] );
        
        dwLcid = *(( LCID __RPC_FAR * )_StubMsg.Buffer)++;
        
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCCommon*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> SetLocaleID((IOPCCommon *) ((CStdStubBuffer *)This)->pvServerObject,dwLcid);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 4U;
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}


HRESULT STDMETHODCALLTYPE IOPCCommon_GetLocaleID_Proxy( 
    IOPCCommon __RPC_FAR * This,
    /* [out] */ LCID __RPC_FAR *pdwLcid)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      4);
        
        
        
        if(!pdwLcid)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 0U;
            NdrProxyGetBuffer(This, &_StubMsg);
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[10] );
            
            *pdwLcid = *(( LCID __RPC_FAR * )_StubMsg.Buffer)++;
            
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        NdrClearOutParameters(
                         ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                         ( PFORMAT_STRING  )&__MIDL_TypeFormatString.Format[6],
                         ( void __RPC_FAR * )pdwLcid);
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCCommon_GetLocaleID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    LCID _M2;
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    LCID __RPC_FAR *pdwLcid;
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    ( LCID __RPC_FAR * )pdwLcid = 0;
    RpcTryFinally
        {
        pdwLcid = &_M2;
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCCommon*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> GetLocaleID((IOPCCommon *) ((CStdStubBuffer *)This)->pvServerObject,pdwLcid);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 4U + 4U;
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        *(( LCID __RPC_FAR * )_StubMsg.Buffer)++ = *pdwLcid;
        
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}


HRESULT STDMETHODCALLTYPE IOPCCommon_QueryAvailableLocaleIDs_Proxy( 
    IOPCCommon __RPC_FAR * This,
    /* [out] */ DWORD __RPC_FAR *pdwCount,
    /* [size_is][size_is][out] */ LCID __RPC_FAR *__RPC_FAR *pdwLcid)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    if(pdwLcid)
        {
        *pdwLcid = 0;
        }
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      5);
        
        
        
        if(!pdwCount)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        if(!pdwLcid)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 0U;
            NdrProxyGetBuffer(This, &_StubMsg);
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[16] );
            
            *pdwCount = *(( DWORD __RPC_FAR * )_StubMsg.Buffer)++;
            
            NdrPointerUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                  (unsigned char __RPC_FAR * __RPC_FAR *)&pdwLcid,
                                  (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[10],
                                  (unsigned char)0 );
            
            _StubMsg.Buffer = (unsigned char __RPC_FAR *)(((long)_StubMsg.Buffer + 3) & ~ 0x3);
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        NdrClearOutParameters(
                         ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                         ( PFORMAT_STRING  )&__MIDL_TypeFormatString.Format[6],
                         ( void __RPC_FAR * )pdwCount);
        _StubMsg.MaxCount = pdwCount ? *pdwCount : 0;
        
        NdrClearOutParameters(
                         ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                         ( PFORMAT_STRING  )&__MIDL_TypeFormatString.Format[10],
                         ( void __RPC_FAR * )pdwLcid);
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCCommon_QueryAvailableLocaleIDs_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    DWORD _M4;
    LCID __RPC_FAR *_M5;
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    DWORD __RPC_FAR *pdwCount;
    LCID __RPC_FAR *__RPC_FAR *pdwLcid;
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    ( DWORD __RPC_FAR * )pdwCount = 0;
    ( LCID __RPC_FAR *__RPC_FAR * )pdwLcid = 0;
    RpcTryFinally
        {
        pdwCount = &_M4;
        pdwLcid = &_M5;
        _M5 = 0;
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCCommon*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> QueryAvailableLocaleIDs(
                           (IOPCCommon *) ((CStdStubBuffer *)This)->pvServerObject,
                           pdwCount,
                           pdwLcid);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 4U + 8U + 7U;
        _StubMsg.MaxCount = pdwCount ? *pdwCount : 0;
        
        NdrPointerBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                              (unsigned char __RPC_FAR *)pdwLcid,
                              (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[10] );
        
        _StubMsg.BufferLength += 16;
        
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        *(( DWORD __RPC_FAR * )_StubMsg.Buffer)++ = *pdwCount;
        
        _StubMsg.MaxCount = pdwCount ? *pdwCount : 0;
        
        NdrPointerMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                            (unsigned char __RPC_FAR *)pdwLcid,
                            (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[10] );
        
        _StubMsg.Buffer = (unsigned char __RPC_FAR *)(((long)_StubMsg.Buffer + 3) & ~ 0x3);
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        _StubMsg.MaxCount = pdwCount ? *pdwCount : 0;
        
        NdrPointerFree( &_StubMsg,
                        (unsigned char __RPC_FAR *)pdwLcid,
                        &__MIDL_TypeFormatString.Format[10] );
        
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}


HRESULT STDMETHODCALLTYPE IOPCCommon_GetErrorString_Proxy( 
    IOPCCommon __RPC_FAR * This,
    /* [in] */ HRESULT dwError,
    /* [string][out] */ LPWSTR __RPC_FAR *ppString)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    if(ppString)
        {
        *ppString = 0;
        }
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      6);
        
        
        
        if(!ppString)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 4U;
            NdrProxyGetBuffer(This, &_StubMsg);
            *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = dwError;
            
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[26] );
            
            NdrPointerUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                  (unsigned char __RPC_FAR * __RPC_FAR *)&ppString,
                                  (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28],
                                  (unsigned char)0 );
            
            _StubMsg.Buffer = (unsigned char __RPC_FAR *)(((long)_StubMsg.Buffer + 3) & ~ 0x3);
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        NdrClearOutParameters(
                         ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                         ( PFORMAT_STRING  )&__MIDL_TypeFormatString.Format[28],
                         ( void __RPC_FAR * )ppString);
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCCommon_GetErrorString_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    LPWSTR _M8;
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    HRESULT dwError;
    LPWSTR __RPC_FAR *ppString;
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    ( LPWSTR __RPC_FAR * )ppString = 0;
    RpcTryFinally
        {
        if ( (_pRpcMessage->DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
            NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[26] );
        
        dwError = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
        
        ppString = &_M8;
        _M8 = 0;
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCCommon*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> GetErrorString(
                  (IOPCCommon *) ((CStdStubBuffer *)This)->pvServerObject,
                  dwError,
                  ppString);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 16U + 10U;
        NdrPointerBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                              (unsigned char __RPC_FAR *)ppString,
                              (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28] );
        
        _StubMsg.BufferLength += 16;
        
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        NdrPointerMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                            (unsigned char __RPC_FAR *)ppString,
                            (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28] );
        
        _StubMsg.Buffer = (unsigned char __RPC_FAR *)(((long)_StubMsg.Buffer + 3) & ~ 0x3);
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        NdrPointerFree( &_StubMsg,
                        (unsigned char __RPC_FAR *)ppString,
                        &__MIDL_TypeFormatString.Format[28] );
        
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}


HRESULT STDMETHODCALLTYPE IOPCCommon_SetClientName_Proxy( 
    IOPCCommon __RPC_FAR * This,
    /* [string][in] */ LPCWSTR szName)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      7);
        
        
        
        if(!szName)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 12U;
            NdrConformantStringBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                           (unsigned char __RPC_FAR *)szName,
                                           (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4] );
            
            NdrProxyGetBuffer(This, &_StubMsg);
            NdrConformantStringMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                                         (unsigned char __RPC_FAR *)szName,
                                         (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4] );
            
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[0] );
            
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCCommon_SetClientName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    LPCWSTR szName;
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    ( LPCWSTR  )szName = 0;
    RpcTryFinally
        {
        if ( (_pRpcMessage->DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
            NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[0] );
        
        NdrConformantStringUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                       (unsigned char __RPC_FAR * __RPC_FAR *)&szName,
                                       (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4],
                                       (unsigned char)0 );
        
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCCommon*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> SetClientName((IOPCCommon *) ((CStdStubBuffer *)This)->pvServerObject,szName);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 4U;
        _StubMsg.BufferLength += 16;
        
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}

const CINTERFACE_PROXY_VTABLE(8) _IOPCCommonProxyVtbl = 
{
    &IID_IOPCCommon,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    IOPCCommon_SetLocaleID_Proxy ,
    IOPCCommon_GetLocaleID_Proxy ,
    IOPCCommon_QueryAvailableLocaleIDs_Proxy ,
    IOPCCommon_GetErrorString_Proxy ,
    IOPCCommon_SetClientName_Proxy
};


static const PRPC_STUB_FUNCTION IOPCCommon_table[] =
{
    IOPCCommon_SetLocaleID_Stub,
    IOPCCommon_GetLocaleID_Stub,
    IOPCCommon_QueryAvailableLocaleIDs_Stub,
    IOPCCommon_GetErrorString_Stub,
    IOPCCommon_SetClientName_Stub
};

const CInterfaceStubVtbl _IOPCCommonStubVtbl =
{
    &IID_IOPCCommon,
    0,
    8,
    &IOPCCommon_table[-3],
    CStdStubBuffer_METHODS
};


/* Object interface: IOPCServerList, ver. 0.0,
   GUID={0x13486D50,0x4821,0x11D2,{0xA4,0x94,0x3C,0xB3,0x06,0xC1,0x00,0x00}} */


extern const MIDL_STUB_DESC Object_StubDesc;


#pragma code_seg(".orpc")

HRESULT STDMETHODCALLTYPE IOPCServerList_EnumClassesOfCategories_Proxy( 
    IOPCServerList __RPC_FAR * This,
    /* [in] */ ULONG cImplemented,
    /* [size_is][in] */ CATID __RPC_FAR rgcatidImpl[  ],
    /* [in] */ ULONG cRequired,
    /* [size_is][in] */ CATID __RPC_FAR rgcatidReq[  ],
    /* [out] */ IEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    if(ppenumClsid)
        {
        MIDL_memset(
               ppenumClsid,
               0,
               sizeof( IEnumGUID __RPC_FAR *__RPC_FAR * ));
        }
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      3);
        
        
        
        if(!ppenumClsid)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 4U + 4U + 7U + 7U;
            _StubMsg.MaxCount = cImplemented;
            
            NdrConformantArrayBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                          (unsigned char __RPC_FAR *)rgcatidImpl,
                                          (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[54] );
            
            _StubMsg.MaxCount = cRequired;
            
            NdrConformantArrayBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                          (unsigned char __RPC_FAR *)rgcatidReq,
                                          (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[68] );
            
            NdrProxyGetBuffer(This, &_StubMsg);
            *(( ULONG __RPC_FAR * )_StubMsg.Buffer)++ = cImplemented;
            
            _StubMsg.MaxCount = cImplemented;
            
            NdrConformantArrayMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                                        (unsigned char __RPC_FAR *)rgcatidImpl,
                                        (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[54] );
            
            *(( ULONG __RPC_FAR * )_StubMsg.Buffer)++ = cRequired;
            
            _StubMsg.MaxCount = cRequired;
            
            NdrConformantArrayMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                                        (unsigned char __RPC_FAR *)rgcatidReq,
                                        (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[68] );
            
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[34] );
            
            NdrPointerUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                  (unsigned char __RPC_FAR * __RPC_FAR *)&ppenumClsid,
                                  (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[82],
                                  (unsigned char)0 );
            
            _StubMsg.Buffer = (unsigned char __RPC_FAR *)(((long)_StubMsg.Buffer + 3) & ~ 0x3);
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        NdrClearOutParameters(
                         ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                         ( PFORMAT_STRING  )&__MIDL_TypeFormatString.Format[82],
                         ( void __RPC_FAR * )ppenumClsid);
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCServerList_EnumClassesOfCategories_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    IEnumGUID __RPC_FAR *_M13;
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    ULONG cImplemented;
    ULONG cRequired;
    IEnumGUID __RPC_FAR *__RPC_FAR *ppenumClsid;
    CATID ( __RPC_FAR *rgcatidImpl )[  ];
    CATID ( __RPC_FAR *rgcatidReq )[  ];
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    rgcatidImpl = 0;
    rgcatidReq = 0;
    ( IEnumGUID __RPC_FAR *__RPC_FAR * )ppenumClsid = 0;
    RpcTryFinally
        {
        if ( (_pRpcMessage->DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
            NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[34] );
        
        cImplemented = *(( ULONG __RPC_FAR * )_StubMsg.Buffer)++;
        
        NdrConformantArrayUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                      (unsigned char __RPC_FAR * __RPC_FAR *)&rgcatidImpl,
                                      (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[54],
                                      (unsigned char)0 );
        
        cRequired = *(( ULONG __RPC_FAR * )_StubMsg.Buffer)++;
        
        NdrConformantArrayUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                      (unsigned char __RPC_FAR * __RPC_FAR *)&rgcatidReq,
                                      (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[68],
                                      (unsigned char)0 );
        
        ppenumClsid = &_M13;
        _M13 = 0;
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCServerList*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> EnumClassesOfCategories(
                           (IOPCServerList *) ((CStdStubBuffer *)This)->pvServerObject,
                           cImplemented,
                           *rgcatidImpl,
                           cRequired,
                           *rgcatidReq,
                           ppenumClsid);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 0U + 4U;
        NdrPointerBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                              (unsigned char __RPC_FAR *)ppenumClsid,
                              (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[82] );
        
        _StubMsg.BufferLength += 16;
        
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        NdrPointerMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                            (unsigned char __RPC_FAR *)ppenumClsid,
                            (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[82] );
        
        _StubMsg.Buffer = (unsigned char __RPC_FAR *)(((long)_StubMsg.Buffer + 3) & ~ 0x3);
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        NdrPointerFree( &_StubMsg,
                        (unsigned char __RPC_FAR *)ppenumClsid,
                        &__MIDL_TypeFormatString.Format[82] );
        
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}


HRESULT STDMETHODCALLTYPE IOPCServerList_GetClassDetails_Proxy( 
    IOPCServerList __RPC_FAR * This,
    /* [in] */ REFCLSID clsid,
    /* [out] */ LPOLESTR __RPC_FAR *ppszProgID,
    /* [out] */ LPOLESTR __RPC_FAR *ppszUserType)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    if(ppszProgID)
        {
        *ppszProgID = 0;
        }
    if(ppszUserType)
        {
        *ppszUserType = 0;
        }
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      4);
        
        
        
        if(!clsid)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        if(!ppszProgID)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        if(!ppszUserType)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 0U;
            NdrSimpleStructBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                       (unsigned char __RPC_FAR *)clsid,
                                       (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[42] );
            
            NdrProxyGetBuffer(This, &_StubMsg);
            NdrSimpleStructMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                                     (unsigned char __RPC_FAR *)clsid,
                                     (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[42] );
            
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[52] );
            
            NdrPointerUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                  (unsigned char __RPC_FAR * __RPC_FAR *)&ppszProgID,
                                  (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28],
                                  (unsigned char)0 );
            
            NdrPointerUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                  (unsigned char __RPC_FAR * __RPC_FAR *)&ppszUserType,
                                  (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28],
                                  (unsigned char)0 );
            
            _StubMsg.Buffer = (unsigned char __RPC_FAR *)(((long)_StubMsg.Buffer + 3) & ~ 0x3);
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        NdrClearOutParameters(
                         ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                         ( PFORMAT_STRING  )&__MIDL_TypeFormatString.Format[28],
                         ( void __RPC_FAR * )ppszProgID);
        NdrClearOutParameters(
                         ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                         ( PFORMAT_STRING  )&__MIDL_TypeFormatString.Format[28],
                         ( void __RPC_FAR * )ppszUserType);
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCServerList_GetClassDetails_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    LPOLESTR _M18;
    LPOLESTR _M19;
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    REFCLSID clsid;
    LPOLESTR __RPC_FAR *ppszProgID;
    LPOLESTR __RPC_FAR *ppszUserType;
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    ( REFCLSID  )clsid = 0;
    ( LPOLESTR __RPC_FAR * )ppszProgID = 0;
    ( LPOLESTR __RPC_FAR * )ppszUserType = 0;
    RpcTryFinally
        {
        if ( (_pRpcMessage->DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
            NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[52] );
        
        NdrSimpleStructUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                   (unsigned char __RPC_FAR * __RPC_FAR *)&clsid,
                                   (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[42],
                                   (unsigned char)0 );
        
        ppszProgID = &_M18;
        _M18 = 0;
        ppszUserType = &_M19;
        _M19 = 0;
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCServerList*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> GetClassDetails(
                   (IOPCServerList *) ((CStdStubBuffer *)This)->pvServerObject,
                   clsid,
                   ppszProgID,
                   ppszUserType);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 16U + 24U + 10U;
        NdrPointerBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                              (unsigned char __RPC_FAR *)ppszProgID,
                              (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28] );
        
        NdrPointerBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                              (unsigned char __RPC_FAR *)ppszUserType,
                              (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28] );
        
        _StubMsg.BufferLength += 16;
        
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        NdrPointerMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                            (unsigned char __RPC_FAR *)ppszProgID,
                            (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28] );
        
        NdrPointerMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                            (unsigned char __RPC_FAR *)ppszUserType,
                            (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[28] );
        
        _StubMsg.Buffer = (unsigned char __RPC_FAR *)(((long)_StubMsg.Buffer + 3) & ~ 0x3);
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        NdrPointerFree( &_StubMsg,
                        (unsigned char __RPC_FAR *)ppszProgID,
                        &__MIDL_TypeFormatString.Format[28] );
        
        NdrPointerFree( &_StubMsg,
                        (unsigned char __RPC_FAR *)ppszUserType,
                        &__MIDL_TypeFormatString.Format[28] );
        
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}


HRESULT STDMETHODCALLTYPE IOPCServerList_CLSIDFromProgID_Proxy( 
    IOPCServerList __RPC_FAR * This,
    /* [in] */ LPCOLESTR szProgId,
    /* [out] */ LPCLSID clsid)
{

    HRESULT _RetVal;
    
    RPC_MESSAGE _RpcMessage;
    
    MIDL_STUB_MESSAGE _StubMsg;
    
    if(clsid)
        {
        MIDL_memset(
               clsid,
               0,
               sizeof( GUID  ));
        }
    RpcTryExcept
        {
        NdrProxyInitialize(
                      ( void __RPC_FAR *  )This,
                      ( PRPC_MESSAGE  )&_RpcMessage,
                      ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                      ( PMIDL_STUB_DESC  )&Object_StubDesc,
                      5);
        
        
        
        if(!szProgId)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        if(!clsid)
            {
            RpcRaiseException(RPC_X_NULL_REF_POINTER);
            }
        RpcTryFinally
            {
            
            _StubMsg.BufferLength = 12U;
            NdrConformantStringBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                           (unsigned char __RPC_FAR *)szProgId,
                                           (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4] );
            
            NdrProxyGetBuffer(This, &_StubMsg);
            NdrConformantStringMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                                         (unsigned char __RPC_FAR *)szProgId,
                                         (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4] );
            
            NdrProxySendReceive(This, &_StubMsg);
            
            if ( (_RpcMessage.DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
                NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[66] );
            
            NdrSimpleStructUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                       (unsigned char __RPC_FAR * __RPC_FAR *)&clsid,
                                       (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[42],
                                       (unsigned char)0 );
            
            _RetVal = *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++;
            
            }
        RpcFinally
            {
            NdrProxyFreeBuffer(This, &_StubMsg);
            
            }
        RpcEndFinally
        
        }
    RpcExcept(_StubMsg.dwStubPhase != PROXY_SENDRECEIVE)
        {
        NdrClearOutParameters(
                         ( PMIDL_STUB_MESSAGE  )&_StubMsg,
                         ( PFORMAT_STRING  )&__MIDL_TypeFormatString.Format[108],
                         ( void __RPC_FAR * )clsid);
        _RetVal = NdrProxyErrorHandler(RpcExceptionCode());
        }
    RpcEndExcept
    return _RetVal;
}

void __RPC_STUB IOPCServerList_CLSIDFromProgID_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase)
{
    HRESULT _RetVal;
    MIDL_STUB_MESSAGE _StubMsg;
    GUID _clsidM;
    LPCLSID clsid;
    LPCOLESTR szProgId;
    
NdrStubInitialize(
                     _pRpcMessage,
                     &_StubMsg,
                     &Object_StubDesc,
                     _pRpcChannelBuffer);
    ( LPCOLESTR  )szProgId = 0;
    ( LPCLSID  )clsid = 0;
    RpcTryFinally
        {
        if ( (_pRpcMessage->DataRepresentation & 0X0000FFFFUL) != NDR_LOCAL_DATA_REPRESENTATION )
            NdrConvert( (PMIDL_STUB_MESSAGE) &_StubMsg, (PFORMAT_STRING) &__MIDL_ProcFormatString.Format[66] );
        
        NdrConformantStringUnmarshall( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                       (unsigned char __RPC_FAR * __RPC_FAR *)&szProgId,
                                       (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[4],
                                       (unsigned char)0 );
        
        clsid = &_clsidM;
        
        *_pdwStubPhase = STUB_CALL_SERVER;
        _RetVal = (((IOPCServerList*) ((CStdStubBuffer *)This)->pvServerObject)->lpVtbl) -> CLSIDFromProgID(
                   (IOPCServerList *) ((CStdStubBuffer *)This)->pvServerObject,
                   szProgId,
                   clsid);
        
        *_pdwStubPhase = STUB_MARSHAL;
        
        _StubMsg.BufferLength = 0U + 11U;
        NdrSimpleStructBufferSize( (PMIDL_STUB_MESSAGE) &_StubMsg,
                                   (unsigned char __RPC_FAR *)clsid,
                                   (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[42] );
        
        _StubMsg.BufferLength += 16;
        
        NdrStubGetBuffer(This, _pRpcChannelBuffer, &_StubMsg);
        NdrSimpleStructMarshall( (PMIDL_STUB_MESSAGE)& _StubMsg,
                                 (unsigned char __RPC_FAR *)clsid,
                                 (PFORMAT_STRING) &__MIDL_TypeFormatString.Format[42] );
        
        *(( HRESULT __RPC_FAR * )_StubMsg.Buffer)++ = _RetVal;
        
        }
    RpcFinally
        {
        }
    RpcEndFinally
    _pRpcMessage->BufferLength = 
        (unsigned int)((long)_StubMsg.Buffer - (long)_pRpcMessage->Buffer);
    
}


static const MIDL_STUB_DESC Object_StubDesc = 
    {
    0,
    NdrOleAllocate,
    NdrOleFree,
    0,
    0,
    0,
    0,
    0,
    __MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x10001, /* Ndr library version */
    0,
    0x50100a4, /* MIDL Version 5.1.164 */
    0,
    0,
    0,  /* notify & notify_flag routine table */
    1,  /* Flags */
    0,  /* Reserved3 */
    0,  /* Reserved4 */
    0   /* Reserved5 */
    };

const CINTERFACE_PROXY_VTABLE(6) _IOPCServerListProxyVtbl = 
{
    &IID_IOPCServerList,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    IOPCServerList_EnumClassesOfCategories_Proxy ,
    IOPCServerList_GetClassDetails_Proxy ,
    IOPCServerList_CLSIDFromProgID_Proxy
};


static const PRPC_STUB_FUNCTION IOPCServerList_table[] =
{
    IOPCServerList_EnumClassesOfCategories_Stub,
    IOPCServerList_GetClassDetails_Stub,
    IOPCServerList_CLSIDFromProgID_Stub
};

const CInterfaceStubVtbl _IOPCServerListStubVtbl =
{
    &IID_IOPCServerList,
    0,
    6,
    &IOPCServerList_table[-3],
    CStdStubBuffer_METHODS
};

#pragma data_seg(".rdata")

#if !defined(__RPC_WIN32__)
#error  Invalid build platform for this stub.
#endif

static const MIDL_PROC_FORMAT_STRING __MIDL_ProcFormatString =
    {
        0,
        {
			
			0x4d,		/* FC_IN_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/*  2 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */
/*  4 */	0x53,		/* FC_RETURN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/*  6 */	0x4e,		/* FC_IN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/*  8 */	0x53,		/* FC_RETURN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 10 */	
			0x51,		/* FC_OUT_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 12 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */
/* 14 */	0x53,		/* FC_RETURN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 16 */	
			0x51,		/* FC_OUT_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 18 */	NdrFcShort( 0x6 ),	/* Type Offset=6 */
/* 20 */	
			0x51,		/* FC_OUT_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 22 */	NdrFcShort( 0xa ),	/* Type Offset=10 */
/* 24 */	0x53,		/* FC_RETURN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 26 */	0x4e,		/* FC_IN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 28 */	
			0x51,		/* FC_OUT_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 30 */	NdrFcShort( 0x1c ),	/* Type Offset=28 */
/* 32 */	0x53,		/* FC_RETURN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 34 */	0x4e,		/* FC_IN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 36 */	
			0x4d,		/* FC_IN_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 38 */	NdrFcShort( 0x36 ),	/* Type Offset=54 */
/* 40 */	0x4e,		/* FC_IN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 42 */	
			0x4d,		/* FC_IN_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 44 */	NdrFcShort( 0x44 ),	/* Type Offset=68 */
/* 46 */	
			0x51,		/* FC_OUT_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 48 */	NdrFcShort( 0x52 ),	/* Type Offset=82 */
/* 50 */	0x53,		/* FC_RETURN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 52 */	
			0x4d,		/* FC_IN_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 54 */	NdrFcShort( 0x68 ),	/* Type Offset=104 */
/* 56 */	
			0x51,		/* FC_OUT_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 58 */	NdrFcShort( 0x1c ),	/* Type Offset=28 */
/* 60 */	
			0x51,		/* FC_OUT_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 62 */	NdrFcShort( 0x1c ),	/* Type Offset=28 */
/* 64 */	0x53,		/* FC_RETURN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */
/* 66 */	
			0x4d,		/* FC_IN_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 68 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */
/* 70 */	
			0x51,		/* FC_OUT_PARAM */
#ifndef _ALPHA_
			0x1,		/* x86, MIPS & PPC Stack size = 1 */
#else
			0x2,		/* Alpha Stack size = 2 */
#endif
/* 72 */	NdrFcShort( 0x6c ),	/* Type Offset=108 */
/* 74 */	0x53,		/* FC_RETURN_PARAM_BASETYPE */
			0x8,		/* FC_LONG */

			0x0
        }
    };

static const MIDL_TYPE_FORMAT_STRING __MIDL_TypeFormatString =
    {
        0,
        {
			NdrFcShort( 0x0 ),	/* 0 */
/*  2 */	
			0x11, 0x8,	/* FC_RP [simple_pointer] */
/*  4 */	
			0x25,		/* FC_C_WSTRING */
			0x5c,		/* FC_PAD */
/*  6 */	
			0x11, 0xc,	/* FC_RP [alloced_on_stack] [simple_pointer] */
/*  8 */	0x8,		/* FC_LONG */
			0x5c,		/* FC_PAD */
/* 10 */	
			0x11, 0x14,	/* FC_RP [alloced_on_stack] */
/* 12 */	NdrFcShort( 0x2 ),	/* Offset= 2 (14) */
/* 14 */	
			0x13, 0x0,	/* FC_OP */
/* 16 */	NdrFcShort( 0x2 ),	/* Offset= 2 (18) */
/* 18 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 20 */	NdrFcShort( 0x4 ),	/* 4 */
/* 22 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x54,		/* FC_DEREFERENCE */
#ifndef _ALPHA_
/* 24 */	NdrFcShort( 0x4 ),	/* x86, MIPS, PPC Stack size/offset = 4 */
#else
			NdrFcShort( 0x8 ),	/* Alpha Stack size/offset = 8 */
#endif
/* 26 */	0x8,		/* FC_LONG */
			0x5b,		/* FC_END */
/* 28 */	
			0x11, 0x14,	/* FC_RP [alloced_on_stack] */
/* 30 */	NdrFcShort( 0x2 ),	/* Offset= 2 (32) */
/* 32 */	
			0x13, 0x8,	/* FC_OP [simple_pointer] */
/* 34 */	
			0x25,		/* FC_C_WSTRING */
			0x5c,		/* FC_PAD */
/* 36 */	
			0x1d,		/* FC_SMFARRAY */
			0x0,		/* 0 */
/* 38 */	NdrFcShort( 0x8 ),	/* 8 */
/* 40 */	0x2,		/* FC_CHAR */
			0x5b,		/* FC_END */
/* 42 */	
			0x15,		/* FC_STRUCT */
			0x3,		/* 3 */
/* 44 */	NdrFcShort( 0x10 ),	/* 16 */
/* 46 */	0x8,		/* FC_LONG */
			0x6,		/* FC_SHORT */
/* 48 */	0x6,		/* FC_SHORT */
			0x4c,		/* FC_EMBEDDED_COMPLEX */
/* 50 */	0x0,		/* 0 */
			NdrFcShort( 0xfffffff1 ),	/* Offset= -15 (36) */
			0x5b,		/* FC_END */
/* 54 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 56 */	NdrFcShort( 0x10 ),	/* 16 */
/* 58 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x0,		/*  */
#ifndef _ALPHA_
/* 60 */	NdrFcShort( 0x4 ),	/* x86, MIPS, PPC Stack size/offset = 4 */
#else
			NdrFcShort( 0x8 ),	/* Alpha Stack size/offset = 8 */
#endif
/* 62 */	0x4c,		/* FC_EMBEDDED_COMPLEX */
			0x0,		/* 0 */
/* 64 */	NdrFcShort( 0xffffffea ),	/* Offset= -22 (42) */
/* 66 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 68 */	
			0x1b,		/* FC_CARRAY */
			0x3,		/* 3 */
/* 70 */	NdrFcShort( 0x10 ),	/* 16 */
/* 72 */	0x29,		/* Corr desc:  parameter, FC_ULONG */
			0x0,		/*  */
#ifndef _ALPHA_
/* 74 */	NdrFcShort( 0xc ),	/* x86, MIPS, PPC Stack size/offset = 12 */
#else
			NdrFcShort( 0x18 ),	/* Alpha Stack size/offset = 24 */
#endif
/* 76 */	0x4c,		/* FC_EMBEDDED_COMPLEX */
			0x0,		/* 0 */
/* 78 */	NdrFcShort( 0xffffffdc ),	/* Offset= -36 (42) */
/* 80 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 82 */	
			0x11, 0x14,	/* FC_RP [alloced_on_stack] */
/* 84 */	NdrFcShort( 0x2 ),	/* Offset= 2 (86) */
/* 86 */	
			0x2f,		/* FC_IP */
			0x5a,		/* FC_CONSTANT_IID */
/* 88 */	NdrFcLong( 0x2e000 ),	/* 188416 */
/* 92 */	NdrFcShort( 0x0 ),	/* 0 */
/* 94 */	NdrFcShort( 0x0 ),	/* 0 */
/* 96 */	0xc0,		/* 192 */
			0x0,		/* 0 */
/* 98 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 100 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 102 */	0x0,		/* 0 */
			0x46,		/* 70 */
/* 104 */	
			0x11, 0x0,	/* FC_RP */
/* 106 */	NdrFcShort( 0xffffffc0 ),	/* Offset= -64 (42) */
/* 108 */	
			0x11, 0x4,	/* FC_RP [alloced_on_stack] */
/* 110 */	NdrFcShort( 0xffffffbc ),	/* Offset= -68 (42) */

			0x0
        }
    };

const CInterfaceProxyVtbl * _opccomn_ProxyVtblList[] = 
{
    ( CInterfaceProxyVtbl *) &_IOPCServerListProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IOPCShutdownProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IOPCCommonProxyVtbl,
    0
};

const CInterfaceStubVtbl * _opccomn_StubVtblList[] = 
{
    ( CInterfaceStubVtbl *) &_IOPCServerListStubVtbl,
    ( CInterfaceStubVtbl *) &_IOPCShutdownStubVtbl,
    ( CInterfaceStubVtbl *) &_IOPCCommonStubVtbl,
    0
};

PCInterfaceName const _opccomn_InterfaceNamesList[] = 
{
    "IOPCServerList",
    "IOPCShutdown",
    "IOPCCommon",
    0
};


#define _opccomn_CHECK_IID(n)	IID_GENERIC_CHECK_IID( _opccomn, pIID, n)

int __stdcall _opccomn_IID_Lookup( const IID * pIID, int * pIndex )
{
    IID_BS_LOOKUP_SETUP

    IID_BS_LOOKUP_INITIAL_TEST( _opccomn, 3, 2 )
    IID_BS_LOOKUP_NEXT_TEST( _opccomn, 1 )
    IID_BS_LOOKUP_RETURN_RESULT( _opccomn, 3, *pIndex )
    
}

const ExtendedProxyFileInfo opccomn_ProxyFileInfo = 
{
    (PCInterfaceProxyVtblList *) & _opccomn_ProxyVtblList,
    (PCInterfaceStubVtblList *) & _opccomn_StubVtblList,
    (const PCInterfaceName * ) & _opccomn_InterfaceNamesList,
    0, // no delegation
    & _opccomn_IID_Lookup, 
    3,
    1,
    0, /* table of [async_uuid] interfaces */
    0, /* Filler1 */
    0, /* Filler2 */
    0  /* Filler3 */
};

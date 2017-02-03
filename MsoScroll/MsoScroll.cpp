#include "stdafx.h"
#include "debugtrace.h"

extern "C" IMAGE_DOS_HEADER __ImageBase;
HHOOK g_mouseHook;
IDispatch* g_pApplication;
//LONG g_AppID;//Excel=0, Word=1
BOOL g_bRecurse;

//
//   FUNCTION: AutoWrap(int, VARIANT*, IDispatch*, LPOLESTR, int,...)
//   PURPOSE: Automation helper function. It simplifies most of the low-level 
//      details involved with using IDispatch directly. Feel free to use it 
//      in your own implementations. One caveat is that if you pass multiple 
//      parameters, they need to be passed in reverse-order.
//   PARAMETERS:
//      * autoType - Could be one of these values: DISPATCH_PROPERTYGET, 
//      DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISPATCH_METHOD.
//      * pvResult - Holds the return value in a VARIANT.
//      * pDisp - The IDispatch interface.
//      * ptName - The property/method name exposed by the interface.
//      * cArgs - The count of the arguments.
//   RETURN VALUE: An HRESULT value indicating whether the function succeeds or not.
//   EXAMPLE: 
//      AutoWrap(DISPATCH_METHOD, NULL, pDisp, L"call", 2, parm[1], parm[0]);
//
HRESULT AutoWrap(int autoType, VARIANT *pvResult, IDispatch *pDisp, LPOLESTR ptName, int cArgs...)
{
	// Begin variable-argument list
	va_list marker;
	va_start(marker, cArgs);

	if (!pDisp) return E_INVALIDARG;

	// Variables used
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;

	// Get DISPID for name passed
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, /*LOCALE_USER_DEFAULT*/ LOCALE_SYSTEM_DEFAULT, &dispID);
	if (FAILED(hr))
	{
//#ifdef _DEBUG
		OutputDebugString(_T("AutoWrap::IDispatch->GetIDsOfNames failed\n"));
		_com_error err(hr);
		OutputDebugString(err.ErrorMessage()); OutputDebugString(_T("\n"));
//#endif//_DEBUG
		return hr;
	}

	// Allocate memory for arguments
	VARIANT *pArgs = new VARIANT[cArgs + 1];
	// Extract arguments...
	for(int i=0; i < cArgs; i++)
	{
		pArgs[i] = va_arg(marker, VARIANT);
	}

	// Build DISPPARAMS
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts
	if (autoType & DISPATCH_PROPERTYPUT)
	{
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call
	hr = pDisp->Invoke(dispID, IID_NULL, /*LOCALE_USER_DEFAULT*/ LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
	if FAILED(hr) 
	{
//#ifdef _DEBUG
		OutputDebugString(_T("AutoWrap::IDispatch->Invoke failed\n"));
		_com_error err(hr);
		OutputDebugString(err.ErrorMessage()); OutputDebugString(_T("\n"));
//#endif//_DEBUG
		return hr;
	}

	// End variable-argument section
	va_end(marker);
	delete[] pArgs;
return hr;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LRESULT CALLBACK MouseHookProc(int nCode, WPARAM wMsg, LPARAM lParam)
{
	if (g_bRecurse || (nCode<0)) return CallNextHookEx(g_mouseHook, nCode, wMsg, lParam);

	if ((wMsg==WM_MOUSEWHEEL) && (HIWORD(GetKeyState(VK_SHIFT)) || HIWORD(GetKeyState(VK_MENU)))) //VK_MENU=ALT key
	{
		//if (g_bRecurse) {DBGTRACE("prevent RECURSION-------\n");return 1;}
		g_bRecurse=TRUE;
		//DBGTRACE("WM_MOUSEWHEEL received\n");
		//DBGTRACE("hwnd=0x%x\n",((LPMOUSEHOOKSTRUCT)lParam)->hwnd);
		//DBGTRACE("WM_MOUSEWHEEL+VK_SHIFT|VK_MENU\n");

		short zDelta=GET_WHEEL_DELTA_WPARAM(((LPMOUSEHOOKSTRUCTEX)lParam)->mouseData);
		VARIANT vtActiveWindow;
		VariantInit(&vtActiveWindow);
		//protected view window has no scroll/page members
		if SUCCEEDED(AutoWrap(DISPATCH_PROPERTYGET, &vtActiveWindow, g_pApplication, L"ActiveWindow", 0))
		{
			//Excel and Word both have these methods (can be called as ActiveWindow.ActivePane.LargeScroll or ActiveWindow.LargeScroll)
			//Word can also use "PageScroll" method
			LPOLESTR pstrMethodName= HIWORD(GetKeyState(VK_MENU)) ? L"LargeScroll" : L"SmallScroll";//scroll by line/cell or by page

			VARIANT vt_, vt1;
			vt_.vt = VT_I4;//VB Long type
			vt1.vt = VT_I4;
			vt_.lVal = 0;
			vt1.lVal = 1;
			
			if HIWORD(GetKeyState(VK_SHIFT)) //horizontal
			{
				if (zDelta < 0) AutoWrap(DISPATCH_METHOD, NULL, vtActiveWindow.pdispVal, pstrMethodName, 4, vt_, vt1, vt_, vt_); //Left, Right, Up, Down (reverse order!)
				else            AutoWrap(DISPATCH_METHOD, NULL, vtActiveWindow.pdispVal, pstrMethodName, 4, vt1, vt_, vt_, vt_);
			}
			else //vertical
			{
				//if (g_AppID) //Word
				//{
				//	if (zDelta < 0) AutoWrap(DISPATCH_METHOD, NULL, vtActiveWindow.pdispVal, L"PageScroll", 2, vt_, vt1);//Up, Down
				//	else            AutoWrap(DISPATCH_METHOD, NULL, vtActiveWindow.pdispVal, L"PageScroll", 2, vt1, vt_);
				//}
				//else //Excel
				{
					if (zDelta < 0) AutoWrap(DISPATCH_METHOD, NULL, vtActiveWindow.pdispVal, pstrMethodName, 4, vt_, vt_, vt_, vt1);
					else            AutoWrap(DISPATCH_METHOD, NULL, vtActiveWindow.pdispVal, pstrMethodName, 4, vt_, vt_, vt1, vt_);
				}
			}
			vtActiveWindow.pdispVal->Release();
		}

		g_bRecurse=FALSE;

		/*We must not pass this message to hook chain if there's no opened workbook, WM_MOUSEWHEEL+Shift crashes Excel.
		Excel (and Word) installs it's own application-level hooks (WH_MSGFILTER, WH_KEYBOARD, WH_CBT).
		This bug is present in Excel 97, 2000, 2002 and 2003. */
		return 1;
	}
return CallNextHookEx(g_mouseHook, nCode, wMsg, lParam);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
STDAPI Connect(IDispatch *pApplication)
{
	if (pApplication==NULL) return E_INVALIDARG;
	HRESULT hr=S_OK;

	_ASSERTE(g_pApplication==NULL);
	if (g_pApplication==NULL)
	{
		//get application name
//		VARIANT vtAppName;
//		VariantInit(&vtAppName);
//		if SUCCEEDED(AutoWrap(DISPATCH_PROPERTYGET, &vtAppName, pApplication, L"Name", 0))
//		{
//OutputDebugString(vtAppName.bstrVal);
//			if(0 == wcscmp(vtAppName.bstrVal, L"Microsoft Word")) g_AppID=1;
//			VariantClear(&vtAppName);
			g_pApplication=pApplication;
			g_pApplication->AddRef();
			DBGTRACE("MsoScroll::Connect\n");
		//}
	}
	else hr=ERROR_ALREADY_ASSIGNED;

	if SUCCEEDED(hr)
	{
		_ASSERTE(g_mouseHook==NULL);
		if (g_mouseHook==NULL)
		{
			g_mouseHook=SetWindowsHookEx(WH_MOUSE, MouseHookProc, (HINSTANCE)&__ImageBase, GetCurrentThreadId());
			DBGTRACE("MsoScroll::SetWindowsHookEx\n");
		}
		else {hr=ERROR_ALREADY_EXISTS; DBGTRACE("ERROR_ALREADY_EXISTS\n");}
	}
return hr;
}
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
STDAPI Disconnect()
{
	//_ASSERTE(g_mouseHook);
	if (g_mouseHook)
	{
		UnhookWindowsHookEx(g_mouseHook);
		g_mouseHook=NULL;
		DBGTRACE("MsoScroll::UnhookWindowsHookEx\n");
	}
	//_ASSERTE(g_pApplication);
	if (g_pApplication)
	{
		g_pApplication->Release();
		g_pApplication=NULL;
		DBGTRACE("MsoScroll::g_pApplication->Release\n");
	}
return S_OK;
}
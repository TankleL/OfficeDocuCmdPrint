// def.h
//
// Author: Tankle L.
// Date: February 19, 2016
//
// Description:     Define some constants, marcos or any useful
//				element for Office-Documents-Cmd-Style-Printer.


#pragma once

#define	SAFE_RELEASE(pObject)		{ if(pObject){pObject->Release(); pObject = nullptr;} }
#define	SAFE_DELETE(pObject)		{ if(pObject){delete pObject; pObject = nullptr;} }
#define	SAFE_DELETE_ARR(pObject)	{ if(pObject){delete[] pObject; pObject = nullptr;} }
#define	_FRETHR(hResult)			{ if(FAILED(hResult)){_ShowErrorText(hr); return hr;} }
#define	_ERET(Val)					{ _ShowErrorText(Val); return Val; }
#define _INSIDE_FUNCTION(Object)	Object

#define PUBLIC
#define	PRIVATE
#define	PROTECTED
// OfficeDocuCmdPrint.h
//
// Author: Tankle L.
// Date: February 19, 2016
//
// Description:     Declare types, functions or classes here.


#pragma once


/// <Types>
//

typedef struct _xml_attribute_item
{
	CComBSTR		name;
	CComBSTR		text;
} XMLATTRIBUTEITEM, *PXMLATTRIBUTEITEM;

typedef std::vector<XMLATTRIBUTEITEM> XMLATTRIBUTES;

typedef struct _xml_unique_node
{
	CComBSTR		fullName;
	CComBSTR		text;
	XMLATTRIBUTES	attributes;
}XMLUNIQUENODE, *PXMLUNIQUENODE;

typedef std::vector<XMLUNIQUENODE>	XMLUNIQUENODELIST;

/// </Types>


/// <Functions>
HRESULT PUBLIC FetchUniqueElementsFromXmlDoc(XMLUNIQUENODELIST& list,
	IXMLDOMDocument* const pXmlDoc);
void PUBLIC Backslash2Slash(char* pText);
void PUBLIC Backslash2Slash(std::string& text);
void PUBLIC Slash2Backslash(char* pText);
void PUBLIC Slash2Backslash(std::string& text);
void PUBLIC	Backslash2Dash(char* pText);
void PUBLIC	Backslash2Dash(std::string& text);
bool PUBLIC JudgeXMLFileFromName(char* pText);
bool PUBLIC JudgeRELSFileFromName(char* pText);
wchar_t* PUBLIC Utf82Unicode(const char* utf, size_t &unicode_len);

HRESULT	_INSIDE_FUNCTION(_GetEachUniqNode(XMLUNIQUENODELIST& list,
	IXMLDOMNode* const pXmlNode, const CComBSTR& path);)
void _INSIDE_FUNCTION(_ShowErrorText(const int& errCode);)
/// </Functions>

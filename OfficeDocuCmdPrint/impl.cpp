// impl.cpp
//     impl. = implementation
//
// Author: Tankle L.
// Date: February 19, 2016
//
// 

#include "stdafx.h"

#include "defs.h"
#include "OfficeDocuCmdPrint.h"

using namespace std;

//
//
HRESULT PUBLIC FetchUniqueElementsFromXmlDoc(XMLUNIQUENODELIST& list,
											IXMLDOMDocument* const pXmlDoc)
{
	HRESULT hr = S_OK;
	IXMLDOMNodeList*	pTopNodeList = nullptr;

	// Get top-layout nodes.
	// Most of time, there should be two nodes existing.
	// And there will NEVER be over than two nodes.
	_FRETHR(hr = pXmlDoc->get_childNodes(&pTopNodeList));

	// Get counts of top-layout nodes.
	long countNodes = 0;
	_FRETHR(hr = pTopNodeList->get_length(&countNodes));

	// Arrange full node name.
	IXMLDOMNode*			pTempNode = nullptr;
	CComBSTR				path("");
	for (long i = 0; i < countNodes; ++i)
	{
		_FRETHR(hr = pTopNodeList->get_item(i, &pTempNode));

		_FRETHR(hr = _GetEachUniqNode(list, pTempNode, path));

		// Release resources
		SAFE_RELEASE(pTempNode);
	}

	// Release resources
	SAFE_RELEASE(pTopNodeList);

	return 0;
}


//
//
void PUBLIC Backslash2Slash(char* pText)
{
	size_t	len = strlen(pText);
	for (size_t i = 0; i < len; ++i)
	{
		if (pText[i] == '\\')
			pText[i] = '/';
	}
}


//
//
void PUBLIC Backslash2Slash(std::string& text)
{
	std::string::iterator	iterC;

	for (iterC = text.begin();
		iterC != text.end();
		++iterC)
	{
		if ((*iterC) == '\\')
			(*iterC) = '/';
	}
}


//
//
void PUBLIC Slash2Backslash(char* pText)
{
	size_t	len = strlen(pText);
	for (size_t i = 0; i < len; ++i)
	{
		if (pText[i] == '/')
			pText[i] = '\\';
	}
}


//
//
void PUBLIC Slash2Backslash(std::string& text)
{
	std::string::iterator	iterC;

	for (iterC = text.begin();
		iterC != text.end();
		++iterC)
	{
		if ((*iterC) == '/')
			(*iterC) = '\\';
	}
}


//
//
void PUBLIC	Backslash2Dash(char* pText)
{
	size_t	len = strlen(pText);
	for (size_t i = 0; i < len; ++i)
	{
		if (pText[i] == '\\')
			pText[i] = '-';
	}
}


//
//
void PUBLIC	Backslash2Dash(std::string& text)
{
	std::string::iterator	iterC;

	for (iterC = text.begin();
		iterC != text.end();
		++iterC)
	{
		if ((*iterC) == '\\')
			(*iterC) = '-';
	}
}


//
//
bool PUBLIC JudgeXMLFileFromName(char* pText)
{
	size_t len = strlen(pText);

	if (len < 5)
		return false;
	if (pText[len - 4] != '.' ||
		pText[len - 3] != 'x' ||
		pText[len - 2] != 'm' ||
		pText[len - 1] != 'l')
		return false;

	return true;
}


//
//
bool PUBLIC JudgeRELSFileFromName(char* pText)
{
	size_t len = strlen(pText);

	if (len < 6)
		return false;
	if (pText[len - 5] != '.' ||
		pText[len - 4] != 'r' ||
		pText[len - 3] != 'e' ||
		pText[len - 2] != 'l' ||
		pText[len - 1] != 's')
		return false;

	return true;
}


// Function: Utf82Unicode
//
// Desciption:
//
// Remarks:
//     *     ! Remember this implementation has memory risk, because of
//        it news a memory space but DOSE NOT delete it. You have to d-
//        elete it manually.
//           IT IS "pwText", the RETURN VALUE.
//
// References
//     Implementation from internet. "CSDN"(BBS)
//     Blog Link:http://blog.csdn.net/infoworld/article/details/12312227
//     Author: "infoworld"
//                                         P.S.: Thanks for your sharing.
wchar_t* PUBLIC Utf82Unicode(const char* utf)
{
	if (!utf || !strlen(utf))
		return nullptr;

	size_t unicode_len;
	int dwUnicodeLen = MultiByteToWideChar(CP_UTF8, 0, utf, -1, NULL, 0);
	size_t num = dwUnicodeLen*sizeof(wchar_t);
	wchar_t *pwText = new wchar_t[num + 1];

	memset(pwText, 0, num);
	MultiByteToWideChar(CP_UTF8, 0, utf, -1, pwText, dwUnicodeLen);
	unicode_len = dwUnicodeLen - 1;
	pwText[num] = L'\0';

	return pwText;
}


// Function: Unicode2Utf8
//
// Desciption:
//
// Remarks:
//     *     ! Remember this implementation has memory risk, because of
//        it news a memory space but DOSE NOT delete it. You have to d-
//        elete it manually.
//           IT IS "szUtf8", the RETURN VALUE.
//
// References
//     Implementation from internet. "CSDN"(BBS)
//     Blog Link:http://blog.csdn.net/infoworld/article/details/12312227
//     Author: "infoworld"
//                                         P.S.: Thanks for your sharing.
char* PUBLIC Unicode2Utf8(const wchar_t* unicode)
{
	if (!unicode)
		return nullptr;

    int len;
    len = WideCharToMultiByte(CP_UTF8, 0, unicode, -1, NULL, 0, NULL, NULL);
    char *szUtf8 = new char[len + 1];

    memset(szUtf8, 0, len + 1);
    WideCharToMultiByte(CP_UTF8, 0, unicode, -1, szUtf8, len, NULL, NULL);

    return szUtf8;
}



//
//
HRESULT	_INSIDE_FUNCTION(_GetEachUniqNode(XMLUNIQUENODELIST& list,
										IXMLDOMNode* const pXmlNode,
										const CComBSTR& path))
{
	HRESULT					hr = S_OK;
	IXMLDOMNodeList*		pChildList = nullptr;
	XMLUNIQUENODE			item;
	VARIANT_BOOL			bHasChild = FALSE;

	// Query thar whether the node has any child nodes.
	_FRETHR(hr = pXmlNode->hasChildNodes(&bHasChild));

	// Get temp name.
	CComBSTR	nodeName;
	_FRETHR(hr = pXmlNode->get_nodeName(&nodeName));

	// RECURSIVE BOUNDARY.
	// The node does not have any child nodes.
	CComBSTR			tempPath;
	XMLATTRIBUTEITEM	tempAttribute;
	IXMLDOMNode*		tempNode;

	long					countofAttributes = 0;
	IXMLDOMNamedNodeMap*	pAtMap = nullptr;
	IXMLDOMNode*			pAttribute = nullptr;
	if (bHasChild == FALSE)
	{
		// Retrieve and save the name.
		item.fullName = path;
		item.fullName.Append(L'/');
		item.fullName.AppendBSTR(nodeName.m_str);

		// Judge that whether the node is a pure TEXT node.
		if (nodeName != "#text")
		{
			// Retrieve and save the attributes.
			_FRETHR(hr = pXmlNode->get_attributes(&pAtMap));
			_FRETHR(hr = pAtMap->get_length(&countofAttributes));

			for (long i = 0; i < countofAttributes; ++i)
			{
				_FRETHR(hr = pAtMap->get_item(i, &pAttribute));

				_FRETHR(hr = pAttribute->get_nodeName(&tempAttribute.name));
				_FRETHR(hr = pAttribute->get_text(&tempAttribute.text));

				SAFE_RELEASE(pAttribute);

				item.attributes.push_back(tempAttribute);
			}

			SAFE_RELEASE(pAtMap);
		}
		else{
			// Retrieve and save the text
			_FRETHR(hr = pXmlNode->get_text(&item.text));
		}

		// Fix into result.
		list.push_back(item);

		return S_OK;
	}
	else // Continue the recursion.
	{
		_FRETHR(hr = pXmlNode->get_childNodes(&pChildList));

		// Get counts of child nodes.
		long countNodes = 0;
		_FRETHR(hr = pChildList->get_length(&countNodes));

		// Append the path.
		tempPath = path;
		tempPath.Append(L'/');
		tempPath.Append(nodeName);

		// Retrieve and save the name.
		item.fullName = tempPath;

		// Retrieve and save the attributes.
		_FRETHR(hr = pXmlNode->get_attributes(&pAtMap));
		_FRETHR(hr = pAtMap->get_length(&countofAttributes));

		for (long i = 0; i < countofAttributes; ++i)
		{
			_FRETHR(hr = pAtMap->get_item(i, &pAttribute));

			_FRETHR(hr = pAttribute->get_nodeName(&tempAttribute.name));
			_FRETHR(hr = pAttribute->get_text(&tempAttribute.text));

			SAFE_RELEASE(pAttribute);

			item.attributes.push_back(tempAttribute);
		}

		SAFE_RELEASE(pAtMap);

		// Fix into result.
		list.push_back(item);

		// Go into the recursion.
		for (long i = 0; i < countNodes; ++i)
		{
			_FRETHR(hr = pChildList->get_item(i, &tempNode));

			_FRETHR(hr = _GetEachUniqNode(list, tempNode, tempPath));

			// Release resources
			tempNode->Release();
			tempNode = nullptr;
		}

		// Release resources
		pChildList->Release();

		return hr;
	}
}




//
//
void _INSIDE_FUNCTION(_ShowErrorText(const int& errCode))
{
	string	text;

	switch (errCode)
	{
	case -10000:
		text = "Error: Wrong input arguments. Code:-10000";
		text += "\n\n\t usage: odcp <office-document> <dest-dir> [<-a>] | <-v>";
		text += "\n\t\t -v : See the version of this routine.";
		text += "\n\t\t -a : Do not output as XML files but as analysis files.";
		text += "\n\n\t version: v0.5";
		text += "\n\t author: Tankle L.\tEmail: Tankle_@hotmail.com";
		text += "\n\t date: Februray 19, 2016";
		break;

	case -10001:
		text = "ERROR: Failed to initialize COM.";
		break;

	case -10002:
		text = "ERROR: Cannot unzip the file. It might be corrupted or imcomplete.";
		break;

	case -10003:
		text = "ERROR: Cannot fetch the zip-file's global infomation.";
		break;

	case -10004:
		text = "ERROR: Cannot read sub-file in this zip-file.";
		break;

	case -10005:
		text = "ERROR: XML file is not xml-well-formed. It might be corrupted.";
		break;

	case -10006:
		text = "ERROR: Cannot create an output file on disk. Check out your system access right";
		break;

	case -10007:
		text = "ERROR: Cannot create an output file on disk. Check out your system access right";
		break;

	default:
		text = "Unknown error occurred.";
	}

	cout << text << endl;
}

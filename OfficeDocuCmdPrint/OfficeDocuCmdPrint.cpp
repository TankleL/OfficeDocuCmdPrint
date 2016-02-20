// OfficeDocuCmdPrint.cpp :
//
// Author: Tankle L.
// Date: February 19, 2016

#include "stdafx.h"

#include "defs.h"
#include "OfficeDocuCmdPrint.h"

using namespace std;

extern std::string textVersion;

int main(int argc, char* argv[])
{
	// Initialize
	if (argc != 3 && argc != 4 && argc != 2)
		_ERET(-10000);

	if (FAILED(CoInitialize(nullptr)))
		_ERET(-10001);


	HRESULT				hr = S_OK;
	IXMLDOMDocument*	pXmlDoc = nullptr;

	_FRETHR(hr = CoCreateInstance(CLSID_DOMDocument60, nullptr, CLSCTX_INPROC_SERVER,
		IID_IXMLDOMDocument, (void**)&pXmlDoc));

	_FRETHR(hr = pXmlDoc->put_async(VARIANT_FALSE));
	_FRETHR(hr = pXmlDoc->put_validateOnParse(VARIANT_FALSE));
	_FRETHR(pXmlDoc->put_resolveExternals(VARIANT_FALSE));
	_FRETHR(pXmlDoc->put_preserveWhiteSpace(VARIANT_TRUE));

	// Vars to use.

	// Unzip mode
	if (argc == 3 || argc == 4)
	{
		// Open a zipped file.
		unzFile	zfile = unzOpen64(argv[1]);
		if (!zfile)
			_ERET(-10002);

		// Query zipped file's global infomation.
		unz_global_info64 zipfileGlobalInfo;
		if (unzGetGlobalInfo64(zfile, &zipfileGlobalInfo) != ZIP_OK)
			_ERET(-10003);

		// Operate each sub-files.
		unz_file_info64	subfileInfo;
		char*			cstrFileName = nullptr;
		char*			pBuf = nullptr;
		int				bytesRead = 0;
		for (ZPOS64_T i = 0; i < zipfileGlobalInfo.number_entry; ++i)
		{
			// Get sub-file's infomation.
			if (unzGetCurrentFileInfo64(zfile, &subfileInfo,
					nullptr, 0, nullptr, 0, nullptr, 0) != ZIP_OK)
				_ERET(-10003);

			// New a memory space for sub-file-name. ! Remeber to delete it manually.
			cstrFileName = new char[subfileInfo.size_filename + 1];
			if (unzGetCurrentFileInfo64(zfile, nullptr,
				cstrFileName, subfileInfo.size_filename, nullptr, 0, nullptr, 0) != ZIP_OK)
				_ERET(-10003);
			cstrFileName[subfileInfo.size_filename] = '\0';

			// Judge XML file
			if (JudgeXMLFileFromName(cstrFileName) == false &&
				JudgeRELSFileFromName(cstrFileName) == false)
			{
				SAFE_DELETE_ARR(cstrFileName);
				unzGoToNextFile(zfile);
				continue;
			}


			// Open sub-file.
			if (unzOpenCurrentFile(zfile) != ZIP_OK)
				_ERET(-10004);

			// New a memory space for buffering the xml-data. ! Remeber to delete it manually.			
			pBuf = new char[subfileInfo.uncompressed_size + 2];
			pBuf[subfileInfo.uncompressed_size] = 0;
			pBuf[subfileInfo.uncompressed_size + 1] = 0;

			
			// Read into buffer
			bytesRead = unzReadCurrentFile(zfile, pBuf, (unsigned int)subfileInfo.uncompressed_size);
			
			// Check out the data actually read.
			//      I think that there must be a better way to implement this work, but
			//   for completing the first usable verion, I chose the simplest way.
			//      Here will possibly occur some bugs in some situations.
			//      Try to fix the potential bugs as you can.
			if (bytesRead != subfileInfo.uncompressed_size)
			{
				// Give up this sub-file.
				// Try to analyze the next one.

				// Release memories.
				SAFE_DELETE_ARR(cstrFileName);
				SAFE_DELETE_ARR(pBuf);

				unzCloseCurrentFile(zfile);
				unzGoToNextFile(zfile);
			}
			else
			{
				// Analysise mode.
				if (argc == 4 && strcmp(argv[3], "-a") == 0)
				{
					wchar_t*	pUnicodeBuf = nullptr;
					size_t		unicodelen = 0;

					pUnicodeBuf = Utf82Unicode(pBuf, unicodelen);
					SAFE_DELETE_ARR(pBuf);

					CComBSTR		XmlCode(pUnicodeBuf);
					VARIANT_BOOL	bSucc = FALSE;
									
					// Load xml.
					_FRETHR(hr = pXmlDoc->loadXML(XmlCode.m_str, &bSucc));	
					if (bSucc == FALSE)
						_ERET(-10005);

					// Get top-level node list.
					IXMLDOMNodeList*	pXmlNodeList = nullptr;
					hr = pXmlDoc->get_childNodes(&pXmlNodeList);

					// Analyze xml code.
					XMLUNIQUENODELIST list;
					_FRETHR(hr = FetchUniqueElementsFromXmlDoc(list, pXmlDoc));

					// Abort xml.
					_FRETHR(hr = pXmlDoc->abort());
					
					// Output as a file.
					string ofilename = argv[2];

					// Check out the slashes.
					Slash2Backslash(cstrFileName);
					Backslash2Dash(cstrFileName);

					// Append name.
					ofilename.append(cstrFileName);
					ofilename.append(".odcp");

					ofstream ofile(ofilename, ios::binary);
					if (ofile.is_open() == false)
						_ERET(-10006);
					
					int i = 0;
					for (auto& pNode : list)
					{
						ofile.write((char*)L" * ", 6);
						ofile.write((char*)pNode.fullName.m_str, 2 * pNode.fullName.Length());
						ofile.write((char*)L" {", 4);

						for (auto& at : pNode.attributes)
						{
							ofile.write((char*)L" ", 2);
							ofile.write((char*)at.name.m_str, at.name.Length()*2);
							ofile.write((char*)L" = \"", 8);
							ofile.write((char*)at.text.m_str, at.text.Length()*2);
							ofile.write((char*)L"\";", 4);
						}

						if (pNode.text.Length() != 0)
						{
							ofile.write((char*)L" } #", 8);
							ofile.write((char*)pNode.text.m_str, 2 * pNode.text.Length());
							ofile.write((char*)L"#\n", 4);
						}
						else
						{
							ofile.write((char*)L" } ##(Null)##\n", 28);
						}

						++i;
					}


					ofile.close();		
				}
				else
				// Directly unzip mode
				if (argc == 3)
				{
				}
				// Wrong input arguments.
				else
					_ERET(-10000);


			}

			// Release memories.
			SAFE_DELETE_ARR(cstrFileName);
			SAFE_DELETE_ARR(pBuf);

			// Close sub-file.
			unzCloseCurrentFile(zfile);
			// Next sub-file.
			unzGoToNextFile(zfile);
		}

		// Close the zipped file.
		unzClose(zfile);
	}
	else
	// Info mode
	if (argc == 2)
	{
		if (strcmp(argv[1], "-v") == 0)
		{
			cout << textVersion << endl;
		}
	}
	return 0;
}

std::string textVersion = "\n\n\t version: v0.5\n\
\t author : Tankle L.\tEmail : Tankle_@hotmail.com\n\
\t date: Februray 19, 2016";

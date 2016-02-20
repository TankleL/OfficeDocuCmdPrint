

#pragma once

#include "targetver.h"

#include <stdio.h>
#include <tchar.h>

#include <iostream>
#include <fstream>
#include <string>
#include <vector>

#include <msxml6.h>

#include <atlbase.h>

#include "minizip-lib\zlib.h"
#include "minizip-lib\zip.h"
#include "minizip-lib\unzip.h"

#pragma comment(lib, "msxml6.lib")

#ifdef _DEBUG
#	pragma comment(lib, "minizip-lib/minizip-lib_debug.lib")
#else
#	pragma comment(lib, "minizip-lib/minizip-lib.lib")
#endif

// TODO:  在此处引用程序需要的其他头文件

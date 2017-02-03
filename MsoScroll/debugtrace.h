#include <tchar.h>
#ifndef _DEBUGTRACE_8A7AC67B_B017_4AAB_BBA7_1EE09FEDC416_
#define _DEBUGTRACE_8A7AC67B_B017_4AAB_BBA7_1EE09FEDC416_
#ifdef _DEBUG

#define DBGTRACE(msg, ...) DbgTrace(0, "", msg, __VA_ARGS__)
#define DBGTRACE2(msg, ...) DbgTrace(__LINE__, __FILE__, msg, __VA_ARGS__)

void DbgTrace(int line, const char* fileName, const char* msg, ...)
{
    va_list args;
    char buffer[256] = { 0 };
	if (line)
	{
		/*_stprintf_s*/sprintf_s(buffer, "%s(%d) : ", (strrchr(fileName, '\\') ? strrchr(fileName, '\\') + 1 : fileName), line);
		OutputDebugStringA(buffer);
	}

    // retrieve the variable arguments
    va_start(args, msg);
    /*_vstprintf_s*/vsprintf_s(buffer, msg, args);
    OutputDebugStringA(buffer);
    va_end(args);
}

#else
#define DBGTRACE(msg, ...)
#define DBGTRACE2(msg, ...)
#endif//_DEBUG
#endif//_DEBUGTRACE_8A7AC67B_B017_4AAB_BBA7_1EE09FEDC416_

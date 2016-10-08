// The MIT License (MIT)

// Copyright (c) 2014 Arts Management Systems Ltd.

// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:

// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

#include "OmnisTools.he"
#include "Static.he"

#include "FileObj.he" // License vars

using namespace OmnisTools;

#define LIBXL_STRING_VERSION "3.07.001"   // Compare with libxl.h hex version

// Define static methods
const static qshort cStaticMethodSetKey = 1900;
const static qshort cStaticMethodLibXLVersion = 1901;

// Parameters for Static Methods
// Columns are:
// 1) Name of Parameter (Resource #)
// 2) Return type (fft value)
// 3) Parameter flags of type EXTD_FLAG_xxxx
// 4) Extended flags.  Documentation states, "Must be 0"
ECOparam cStaticMethodsParamsTable[] = 
{
	1950, fftCharacter  , 0, 0,
	1951, fftCharacter  , 0, 0
};

// Table of Methods available for Simple
// Columns are:
// 1) Unique ID 
// 2) Name of Method (Resource #)
// 3) Return Type 
// 4) # of Parameters
// 5) Array of Parameter Names (Taken from MethodsParamsTable.  Increments # of parameters past this pointer) 
// 6) Enum Start (Not sure what this does, 0 = disabled)
// 7) Enum Stop (Not sure what this does, 0 = disabled)
ECOmethodEvent cStaticMethodsTable[] = 
{
	cStaticMethodSetKey, cStaticMethodSetKey, fftNone, 2, &cStaticMethodsParamsTable[0], 0, 0,
	cStaticMethodLibXLVersion, cStaticMethodLibXLVersion, fftNone, 0, 0, 0, 0
};

// List of methods in Simple
qlong returnStaticMethods(tThreadData* pThreadData)
{
	const qshort cStaticMethodCount = sizeof(cStaticMethodsTable) / sizeof(ECOmethodEvent);
	
	return ECOreturnMethods( gInstLib, pThreadData->mEci, &cStaticMethodsTable[0], cStaticMethodCount );
}

// Implementation of single static method
void methodSetKey(tThreadData* pThreadData, qshort paramCount) {
	
	EXTfldval nameVal, keyVal;
	
	if ( getParamVar(pThreadData, 1, nameVal) != qtrue )
		return;
	
	if ( getParamVar(pThreadData, 2, keyVal) != qtrue )
		return;
	
	
	NVObjWorkbook::licenseName = getWStringFromEXTFldVal(nameVal);
	NVObjWorkbook::licenseKey = getWStringFromEXTFldVal(keyVal);
	
	return;
}

// Implementation of single static method
void methodLibXLVersion(tThreadData* pThreadData, qshort paramCount) {	
	
    EXTfldval fValReturn;

	getEXTFldValFromString(fValReturn, LIBXL_STRING_VERSION);

    ECOaddParam(pThreadData->mEci, &fValReturn);

	return;
}

// Static method dispatch
qlong staticMethodCall( OmnisTools::tThreadData* pThreadData ) {
	
	qshort funcId = (qshort)ECOgetId(pThreadData->mEci);
	qshort paramCount = ECOgetParamCount(pThreadData->mEci);
	
	switch( funcId )
	{
		case cStaticMethodSetKey:
			pThreadData->mCurMethodName = "$setKey";
			methodSetKey(pThreadData, paramCount);
			break;

		case cStaticMethodLibXLVersion:
			pThreadData->mCurMethodName = "$libXLVersion";
			methodLibXLVersion(pThreadData, paramCount);
			break;
	}
	
	return 0L;
}


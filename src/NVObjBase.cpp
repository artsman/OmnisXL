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

#include "NVObjBase.he"

using namespace OmnisTools;

// Constructor for building from an existing object instance
NVObjBase::NVObjBase( qobjinst qo ) 
{
	mObjInst = qo;	// we remember the objects instance
}

// Generic destruction
NVObjBase::~NVObjBase()
{
	// Insert any memory deletion code or general cleanup code
}

// Duplicate an object into a new object
NVObjBase* NVObjBase::dup( qlong propID, qobjinst qo, tThreadData *pThreadData  )
{
	NVObjBase* copy = (NVObjBase*)createObject(propID, qo, pThreadData);  // Defined in jsoncpp
	return copy;
}

// Copy one objects contents into another objects contents.
void NVObjBase::copy( NVObjBase* pObj )
{
	qobjinst inst = mObjInst;
	*this = *pObj;
	
	mObjInst = inst;
    mErrorInst = pObj->mErrorInst;
}

// Methods Available and Method Call Handling (These should be overriden by a sub-class)
qlong NVObjBase::returnMethods( tThreadData* pThreadData ) { return 1L; }
qlong NVObjBase::methodCall( tThreadData* pThreadData ) { return 1L; }

// Properties and Property Call Handling (These should be overriden by a sub-class)
qlong NVObjBase::returnProperties( tThreadData* pThreadData ) { return 1L; }
qlong NVObjBase::getProperty( tThreadData* pThreadData ) { return 1L; }
qlong NVObjBase::setProperty( tThreadData* pThreadData ) { return 1L; }
qlong NVObjBase::canAssignProperty( tThreadData* pThreadData, qlong propID ) { return qfalse; }

// Pointer used for sending errors to another instance.
void NVObjBase::setErrorInstance( qobjinst oi ) {
	mErrorInst = oi;
}

void NVObjBase::callErrorMethod( tThreadData* pThreadData, tResult pError )
{
	// If the method completed ok, then there is no need to check for errors.
	if (pError == ERR_OK || pError == ERR_RETURN_ZERO || pError == METHOD_DONE_RETURN || pError == METHOD_OK)
		return;
	
	std::string txt = translateObjectError(pError);
	
	// Error wasn't object specific, look for default error
	if (txt.empty())
		txt = translateDefaultError(pError);
	
	// Fill parameters for error method (ErrorCode, ErrorDesc, ExtraErrorText, MethodName)
	EXTfldval params[4];
	params[0].setLong( pError );
	
	getEXTFldValFromString(params[1], txt);
	
	if ( !(pThreadData->mExtraErrorText.empty()) )
		getEXTFldValFromString(params[2], pThreadData->mExtraErrorText);
	
	if ( !(pThreadData->mCurMethodName.empty()) )
		getEXTFldValFromString(params[3], pThreadData->mCurMethodName);
	
	// Select which instance to use for errors
	qobjinst errInst = (mErrorInst) ? mErrorInst : mObjInst;
	
	// Call $error method
	str31 s_error(initStr31("$error"));
	// DOMETHOD_EXEC_QUEUE is undocumented. It queues the method call so it doesn't interrupt running code.
	ECOdoMethod( errInst, &s_error, params, 4); 
}

// Built in error codes
std::string NVObjBase::translateDefaultError( OmnisTools::tResult pError ) {
	
	std::string txt;
	switch (pError)
	{
		txt = "";
		case ERR_OK:
		case ERR_RETURN_ZERO:
		case METHOD_DONE_RETURN:
			break;
		case ERR_METHOD_FAILED:
			txt = "Method failed.";
			break;
		case ERR_BAD_METHOD:
			txt = "Invalid method called. Internal method index error.";
			break;
		case ERR_BAD_PARAMS:
			txt = "Invalid parameters";
			break;
		case ERR_NOMEMORY:
			txt = "Out of memory.";
			break;
		case ERR_NOT_SUPPORTED:
			txt = "The feature is not supported";
			break;
		case ERR_BAD_CALCULATION:
			txt = "Calculation is invalid";
			break;
		case ERR_NOT_IMPLEMENTED:
			txt = "Feature not implemented.";
			break;
		default:
			break;
	}
	
	return txt;
}

// Object specific error codes (Should be overridden when object specific error codes are specified)
std::string NVObjBase::translateObjectError( OmnisTools::tResult pError ) { 
	return std::string(""); 
}

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

#include "FileObj.he"
#include "WorksheetObj.he" // For reading NVObjWorksheet objects
#include "FormatObj.he" // For reading NVObjFormat objects
#include "FontObj.he" // For reading NVObjFont objects
#include "Constants.he"
#include "OmnisTools.he"

#include <string>
#include <iostream>
#include <map>
#include <boost/format.hpp>

using namespace OmnisTools;
using namespace LibXLConstants;
using namespace libxl;
using boost::shared_ptr;
using boost::format;

// Static variables
std::wstring NVObjWorkbook::licenseName;
std::wstring NVObjWorkbook::licenseKey;

/**************************************************************************************************
 **                       CONSTRUCTORS / DESTRUCTORS                                             **
 **************************************************************************************************/

NVObjWorkbook::NVObjWorkbook(qobjinst qo, tThreadData *pThreadData) : NVObjBase(qo), fileType(XLS_FORMAT)
{ 
	initialize();
}

NVObjWorkbook::~NVObjWorkbook()
{ }

void NVObjWorkbook::copy( NVObjWorkbook* pObj ) {
	// Copy in super class (This does *this = *pObj)
	NVObjBase::copy(pObj);
	
	// Copy the contents of the object into the new object
	book = pObj->book;
	fileType = pObj->fileType;
}

/**************************************************************************************************
 **                              PROPERTY DECLERATION                                            **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project it is also used as the Unique ID.
const static qshort cWBPropertyErrorMessage    = 2500,
					cWBPropertySheetCount      = 2501,
					cWBPropertyFormatSize      = 2502,
					cWBPropertyFontSize        = 2503,
					cWBPropertyActiveSheet     = 2504,
					cWBPropertyDefaultFont     = 2505,
                    cWBPropertyDefaultFontSize = 2506,
                    cWBPropertyRGBMode         = 2507;

/**************************************************************************************************
 **                               METHOD DECLERATION                                             **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project is also used as the Unique ID.
const static qshort cWBMethodInitialize         = 2000,
                    cWBMethodLoad               = 2001,
					cWBMethodSave               = 2002,
					cWBMethodLoadRaw            = 2003,
					cWBMethodSaveRaw            = 2004,
					cWBMethodAddSheet           = 2005,
					cWBMethodGetSheet           = 2006,
					cWBMethodDelSheet           = 2007,
					cWBMethodAddFormat          = 2008,
					cWBMethodAddFont            = 2009,
					cWBMethodAddCustomNumFormat = 2010,
					cWBMethodCustomNumFormat    = 2011,
					cWBMethodFormat             = 2012,
					cWBMethodFont               = 2013,
					cWBMethodAddPicture         = 2014,
					cWBMethodAddPictureRaw      = 2015,
                    cWBMethodPopulateFromList   = 2016,
                    cWBMethodExtractList        = 2017,
                    cWBMethodError              = 2018;

/**************************************************************************************************
 **                                 INSTANCE METHODS                                             **
 **************************************************************************************************/

// Call a method
qlong NVObjWorkbook::methodCall( tThreadData* pThreadData )
{
	if (!book) {
		pThreadData->mExtraErrorText = "Internal Error.  Object dependencies are not setup.";
		callErrorMethod( pThreadData, ERR_METHOD_FAILED );
		return qfalse;
	}
	
	tResult result = ERR_OK;
	qshort funcId = (qshort)ECOgetId(pThreadData->mEci);
	qshort paramCount = ECOgetParamCount(pThreadData->mEci);

	switch( funcId )
	{
		case cWBMethodError:
			result = METHOD_DONE_RETURN;
			break;
		case cWBMethodInitialize:
			pThreadData->mCurMethodName = "$initialize";
			result = methodInitialize(pThreadData, paramCount);
			break;
		case cWBMethodLoad:
			pThreadData->mCurMethodName = "$load";
			result = methodLoad(pThreadData, paramCount);
			break;
		case cWBMethodSave:
			pThreadData->mCurMethodName = "$save";
			result = methodSave(pThreadData, paramCount);
			break;
		case cWBMethodLoadRaw:
			pThreadData->mCurMethodName = "$loadraw";
			result = methodLoadRaw(pThreadData, paramCount);
			break;
		case cWBMethodSaveRaw:
			pThreadData->mCurMethodName = "$saveraw";
			result = methodSaveRaw(pThreadData, paramCount);
			break;
		case cWBMethodAddSheet:
			pThreadData->mCurMethodName = "$addSheet";
			result = methodAddSheet(pThreadData, paramCount);
			break;
		case cWBMethodGetSheet:
			pThreadData->mCurMethodName = "$getSheet";
			result = methodGetSheet(pThreadData, paramCount);
			break;
		case cWBMethodDelSheet:
			pThreadData->mCurMethodName = "$delSheet";
			result = methodDelSheet(pThreadData, paramCount);
			break;
		case cWBMethodAddFormat:
			pThreadData->mCurMethodName = "$addFormat";
			result = methodAddFormat(pThreadData, paramCount);
			break;
		case cWBMethodAddFont:
			pThreadData->mCurMethodName = "$addFont";
			result = methodAddFont(pThreadData, paramCount);
			break;
		case cWBMethodAddCustomNumFormat:
			pThreadData->mCurMethodName = "$addCustomNumFormat";
			result = methodAddCustomNumFormat(pThreadData, paramCount);
			break;
		case cWBMethodCustomNumFormat:
			pThreadData->mCurMethodName = "$customNumFormat";
			result = methodCustomNumFormat(pThreadData, paramCount);
			break;
		case cWBMethodFormat:
			pThreadData->mCurMethodName = "$format";
			result = methodFormat(pThreadData, paramCount);
			break;
		case cWBMethodFont:
			pThreadData->mCurMethodName = "$font";
			result = methodFont(pThreadData, paramCount);
			break;
		case cWBMethodAddPicture:
			pThreadData->mCurMethodName = "$addPicture";
			result = methodAddPicture(pThreadData, paramCount);
			break;
		case cWBMethodAddPictureRaw:
			pThreadData->mCurMethodName = "$addPictureRaw";
			result = methodAddPictureRaw(pThreadData, paramCount);
			break;
		case cWBMethodPopulateFromList:
			pThreadData->mCurMethodName = "$populateFromList";
			result = methodPopulateFromList(pThreadData, paramCount);
			break;
		case cWBMethodExtractList:
			pThreadData->mCurMethodName = "$extractList";
			result = methodExtractList(pThreadData, paramCount);
			break;
		default:
			result = ERR_BAD_METHOD;
			break;
	}
	
	// Check for errors to notify Omnis developer in $error
	callErrorMethod( pThreadData, result );

	return 0L;
}

// Custom errors for object
std::string NVObjWorkbook::translateObjectError( OmnisTools::tResult pError ) {
	
	return std::string("");
}

/**************************************************************************************************
 **                                PROPERTIES                                                    **
 **************************************************************************************************/

// Assignability of properties
qlong NVObjWorkbook::canAssignProperty( tThreadData* pThreadData, qlong propID ) {
	switch (propID) {
		case cWBPropertyErrorMessage:
		case cWBPropertySheetCount:
		case cWBPropertyFormatSize:
		case cWBPropertyFontSize:
			return qfalse;
			break;
		case cWBPropertyActiveSheet:
		case cWBPropertyRGBMode:
		case cWBPropertyDefaultFont:
		case cWBPropertyDefaultFontSize:
			return qtrue;
			break;
		default:
			return qfalse;
	}
}

// Method to retrieve a property of the object
qlong NVObjWorkbook::getProperty( tThreadData* pThreadData ) 
{
	EXTfldval fValReturn;
	if (!book) {
		ECOaddParam(pThreadData->mEci, &fValReturn);
		return qtrue;
	}
	
	int fontSize; // For cWBPropertyDefaultFontSize
	qlong longConv;
	std::string error;
	const wchar_t* valString;
	
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cWBPropertyErrorMessage:
			getEXTFldValFromChar(fValReturn, book->errorMessage());
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cWBPropertySheetCount:
			longConv = static_cast<qlong>(book->sheetCount());
			fValReturn.setLong(longConv); 
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cWBPropertyFormatSize:
			longConv = static_cast<qlong>(book->formatSize());
			fValReturn.setLong(longConv); 
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cWBPropertyFontSize:
			longConv = static_cast<qlong>(book->fontSize());
			fValReturn.setLong(longConv); 
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cWBPropertyActiveSheet:
			if (book->sheetCount() > 0) {
				longConv = static_cast<qlong>(book->activeSheet() + 1);
			} else {
				longConv = 0;
			}
			fValReturn.setLong(longConv); 
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cWBPropertyDefaultFont:
			getEXTFldValFromWChar(fValReturn, book->defaultFont(&fontSize) );
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		
		case cWBPropertyDefaultFontSize:
			valString = book->defaultFont(&fontSize);
			fValReturn.setLong(static_cast<qlong>(fontSize));
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cWBPropertyRGBMode:
			fValReturn.setBool(getQBoolFromBool(book->rgbMode()));
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
	}

	return 1L;
}

// Method to set a property of the object
qlong NVObjWorkbook::setProperty( tThreadData* pThreadData )
{
	if (!book) {
		return qfalse;
	}
	
	// Retrieve value to set for property, always in first parameter
	EXTfldval fVal;
	if( getParamVar( pThreadData->mEci, 1, fVal) == qfalse ) 
		return qfalse;
	
	// Variables needed for type conversion
	int fontSize, activeSheet, test;
	bool rgbBool, rgbTest;
	qbool rgbMode;
	const wchar_t *fontName;
    std::wstring newName;
	
	// Assign to the appropriate property
	// All assignments are protected so that libXL is only called when something changes.
	// This is important since Omnis "sets" variables repeatedly when viewed by instance viewer.
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cWBPropertyActiveSheet:
			activeSheet = static_cast<int>(fVal.getLong() - 1); // LibXL is offset from 0
			test = book->activeSheet();
			if (test != activeSheet) {
				if (activeSheet < book->sheetCount()) {
					book->setActiveSheet(activeSheet);
				}
			}
			break;
		case cWBPropertyDefaultFont:
			newName = getWStringFromEXTFldVal(fVal);
			fontName = book->defaultFont(&fontSize);
			if (newName != fontName) {
				book->setDefaultFont(newName.c_str(), fontSize);
			}
			break;
		case cWBPropertyDefaultFontSize:
			fontName = book->defaultFont(&test);
			fontSize = static_cast<int>(fVal.getLong());
			if (fontSize != test) {
				book->setDefaultFont(fontName, fontSize);
			}
			break;
		case cWBPropertyRGBMode:
			rgbTest = book->rgbMode();
			fVal.getBool(&rgbMode);
			rgbBool = getBoolFromQBool(rgbMode);	
			if (rgbTest != rgbBool && !(rgbBool == true && fileType == XLS_FORMAT)) { // RGB Mode is not allowed for XLS files, just XLSX
				book->setRgbMode(rgbBool);
			}
			break;
	}

	return 1L;
}

/**************************************************************************************************
 **                                        STATIC METHODS                                        **
 **************************************************************************************************/

/* METHODS */

// Table of parameter resources and types.
// Note that all parameters can be stored in this single table and the array offset can be  
// passed via the MethodsTable.
//
// Columns are:
// 1) Name of Parameter (Resource #)
// 2) Return type (fft value)
// 3) Parameter flags of type EXTD_FLAG_xxxx
// 4) Extended flags.  Documentation states, "Must be 0"
ECOparam cWorkbookMethodsParamsTable[] = 
{
	// $initialize
	2900, fftInteger,   EXTD_FLAG_PARAMOPT, 0,
	// $load
	2901, fftCharacter, 0, 0,
	// $save
	2902, fftCharacter, 0, 0,
	//$loadRaw
	2903, fftBinary,    0, 0,
	2904, fftNumber,    0, 0,
	//$saveRaw
	2905, fftBinary,    EXTD_FLAG_PARAMALTER, 0,
	2906, fftNumber,    EXTD_FLAG_PARAMALTER, 0,
	//$addSheet
	2907, fftCharacter, 0, 0,
	2908, fftObject,    0, 0,
	//$getSheet
	2909, fftInteger,   0, 0,
	//$delSheet
	2910, fftInteger,   0, 0,
	//$addFormat
	2911, fftObject,    0, 0,
	//$addFont
	2912, fftObject,    0, 0,
	//$addCustomNumFormat
	2913, fftCharacter, 0, 0,
	//$customNumFormat
	2914, fftInteger,   0, 0,
	//$format
	2915, fftInteger,   0, 0,
	//$font
	2916, fftInteger,   0, 0,
	//$addPicture
	2917, fftCharacter, 0, 0,
	//$addPictureRaw
	2918, fftBinary,    0, 0,
	2919, fftNumber,    0, 0,
	//$populateFromList
	2920, fftList,      0, 0,
	2921, fftInteger,   EXTD_FLAG_PARAMOPT, 0,
	2922, fftRow,       EXTD_FLAG_PARAMOPT, 0,
    2923, fftRow,       EXTD_FLAG_PARAMOPT, 0,
	//$extractList
	2924, fftInteger,   EXTD_FLAG_PARAMOPT, 0,
	2925, fftRow,       EXTD_FLAG_PARAMOPT, 0,
	//$error
	2926, fftInteger,   0, 0,
	2927, fftCharacter, 0, 0,
	2928, fftCharacter, 0, 0,
	2929, fftCharacter, 0, 0
};

// Table of Methods available
// Columns are:
// 1) Unique ID 
// 2) Name of Method (Resource #)
// 3) Return Type 
// 4) # of Parameters
// 5) Array of Parameter Names (Taken from MethodsParamsTable.  Increments # of parameters past this pointer) 
// 6) Enum Start (Not sure what this does, 0 = disabled)
// 7) Enum Stop (Not sure what this does, 0 = disabled)
ECOmethodEvent cWorkbookMethodsTable[] = 
{
	cWBMethodInitialize,         cWBMethodInitialize,         fftNone,      1, &cWorkbookMethodsParamsTable[0],  0, 0,
	cWBMethodLoad,               cWBMethodLoad,               fftBoolean,   1, &cWorkbookMethodsParamsTable[1],  0, 0,
	cWBMethodSave,               cWBMethodSave,               fftBoolean,   1, &cWorkbookMethodsParamsTable[2],  0, 0,
	cWBMethodLoadRaw,            cWBMethodLoadRaw,            fftBoolean,   2, &cWorkbookMethodsParamsTable[3],  0, 0,
	cWBMethodSaveRaw,            cWBMethodSaveRaw,            fftBoolean,   2, &cWorkbookMethodsParamsTable[5],  0, 0,
	cWBMethodAddSheet,           cWBMethodAddSheet,           fftObject,    2, &cWorkbookMethodsParamsTable[7],  0, 0,
	cWBMethodGetSheet,           cWBMethodGetSheet,           fftObject,    1, &cWorkbookMethodsParamsTable[9],  0, 0,
	cWBMethodDelSheet,           cWBMethodDelSheet,           fftBoolean,   1, &cWorkbookMethodsParamsTable[10], 0, 0,
	cWBMethodAddFormat,          cWBMethodAddFormat,          fftObject,    1, &cWorkbookMethodsParamsTable[11], 0, 0,
	cWBMethodAddFont,            cWBMethodAddFont,            fftObject,    1, &cWorkbookMethodsParamsTable[12], 0, 0,
	cWBMethodAddCustomNumFormat, cWBMethodAddCustomNumFormat, fftInteger,   1, &cWorkbookMethodsParamsTable[13], 0, 0,
	cWBMethodCustomNumFormat,    cWBMethodCustomNumFormat,    fftCharacter, 1, &cWorkbookMethodsParamsTable[14], 0, 0,
	cWBMethodFormat,             cWBMethodFormat,             fftObject,    1, &cWorkbookMethodsParamsTable[15], 0, 0,
	cWBMethodFont,               cWBMethodFont,               fftObject,    1, &cWorkbookMethodsParamsTable[16], 0, 0,
	cWBMethodAddPicture,         cWBMethodAddPicture,         fftInteger,   1, &cWorkbookMethodsParamsTable[17], 0, 0,
	cWBMethodAddPictureRaw,      cWBMethodAddPictureRaw,      fftInteger,   2, &cWorkbookMethodsParamsTable[18], 0, 0,
	cWBMethodPopulateFromList,   cWBMethodPopulateFromList,   fftBoolean,   4, &cWorkbookMethodsParamsTable[20], 0, 0,
	cWBMethodExtractList,        cWBMethodExtractList,        fftList,      2, &cWorkbookMethodsParamsTable[24], 0, 0,
	cWBMethodError,              cWBMethodError,              fftNone,      4, &cWorkbookMethodsParamsTable[26], 0, 0
};

// List of methods
qlong NVObjWorkbook::returnMethods(tThreadData* pThreadData)
{
	const qshort cWBMethodCount = sizeof(cWorkbookMethodsTable) / sizeof(ECOmethodEvent);
	
	return ECOreturnMethods( gInstLib, pThreadData->mEci, &cWorkbookMethodsTable[0], cWBMethodCount );
}

/* PROPERTIES */

// Table of properties available
// Columns are:
// 1) Unique ID 
// 2) Name of Property (Resource #)
// 3) Return Type 
// 4) Flags describing the property
// 5) Additional Flags describing the property
// 6) Enum Start (Not sure what this does, 0 = disabled)
// 7) Enum Stop (Not sure what this does, 0 = disabled)
ECOproperty cWorkbookPropertyTable[] = 
{
	cWBPropertyErrorMessage,    cWBPropertyErrorMessage,    fftCharacter, EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cWBPropertySheetCount,      cWBPropertySheetCount,      fftInteger,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cWBPropertyFormatSize,      cWBPropertyFormatSize,      fftInteger,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cWBPropertyFontSize,        cWBPropertyFontSize,        fftInteger,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cWBPropertyActiveSheet,     cWBPropertyActiveSheet,     fftInteger,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cWBPropertyDefaultFont,     cWBPropertyDefaultFont,     fftCharacter, EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cWBPropertyDefaultFontSize, cWBPropertyDefaultFontSize, fftInteger,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cWBPropertyRGBMode,         cWBPropertyRGBMode,         fftBoolean,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0
};

// List of properties 
qlong NVObjWorkbook::returnProperties( tThreadData* pThreadData )
{
	const qshort propertyCount = sizeof(cWorkbookPropertyTable) / sizeof(ECOproperty);

	return ECOreturnProperties( gInstLib, pThreadData->mEci, &cWorkbookPropertyTable[0], propertyCount );
}

/**************************************************************************************************
 **                              OMNIS VISIBLE METHODS                                           **
 **************************************************************************************************/

// Initialize the object to a new file type
OmnisTools::tResult NVObjWorkbook::methodInitialize( tThreadData* pThreadData, qshort pParamCount )
{ 	
	qlong bookParam = XLS_FORMAT;
	getParamLong(pThreadData, 1, bookParam);
	
	fileType = (bookParam == XLSX_FORMAT) ? XLSX_FORMAT : XLS_FORMAT;

	initialize();
	
	return METHOD_DONE_RETURN;
}

// Load the excel file from a given path
OmnisTools::tResult NVObjWorkbook::methodLoad( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	EXTfldval pathVal, returnVal;
	
	if ( getParamVar(pThreadData, 1, pathVal) != qtrue )
		return ERR_BAD_PARAMS;
	
	if ( ensurePosixPath(pathVal) != qtrue ) {
		pThreadData->mExtraErrorText = str(format("Unable to convert HFS Path to Posix: \"%s\"") % getStringFromEXTFldVal(pathVal));
		return ERR_BAD_PARAMS;
	}
	
    initialize(); // Get a new book object before any load so there are no hanging formats and fonts on the old book
    
	std::wstring path = getWStringFromEXTFldVal(pathVal);
	bool success = book->load(path.c_str());
	
	getEXTFldValFromBool(returnVal, success);
	ECOaddParam(pThreadData->mEci, &returnVal);
	
	return METHOD_DONE_RETURN;
}

// Save the Excel file to a given path
OmnisTools::tResult NVObjWorkbook::methodSave( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	EXTfldval pathVal, posixPath, returnVal;
	
	if ( getParamVar(pThreadData, 1, pathVal) != qtrue )
		return ERR_BAD_PARAMS;
	
	if ( ensurePosixPath(pathVal) != qtrue ) {
		pThreadData->mExtraErrorText = str(format("Unable to convert HFS Path to Posix: \"%s\"") % getStringFromEXTFldVal(pathVal));
		return ERR_BAD_PARAMS;
	}
	
	std::wstring path = getWStringFromEXTFldVal(pathVal);
	
	bool success = book->save(path.c_str());
	
	getEXTFldValFromBool(returnVal, success);
	ECOaddParam(pThreadData->mEci, &returnVal);
	
	return METHOD_DONE_RETURN;
}

// Creates a Book from Omnis binary data
OmnisTools::tResult NVObjWorkbook::methodLoadRaw( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	EXTfldval binVal, retVal;
	
	if ( getParamVar(pThreadData, 1, binVal) != qtrue )
		return ERR_BAD_PARAMS;
	
	bool success;
	
	unsigned size;
	qlong omSize = binVal.getBinLen();
	if (omSize == 0) {
		pThreadData->mExtraErrorText = "No data in parameter";
		return ERR_BAD_PARAMS;
	}
	
    initialize(); // Get a new book object before any load so there are no hanging formats and fonts on the old book
    
	const char* data;
	qbyte* omData = static_cast<qbyte*>(MEMmalloc(omSize));
	
	binVal.getBinary(omSize, omData, omSize);
	
	size = static_cast<unsigned>(omSize);
	data = reinterpret_cast<const char*> (omData);
	
	success = book->loadRaw(data, size);
	
	getEXTFldValFromBool(retVal, success);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Creates Omnis binary data from a Book
OmnisTools::tResult NVObjWorkbook::methodSaveRaw( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	EXTfldval binVal, retVal;
	
	if ( getParamVar(pThreadData, 1, binVal) != qtrue )
		return ERR_BAD_PARAMS;
	
	bool success;
	
	const char* data;
	qbyte* omData;
	
	unsigned size;
	qlong omSize;
	
	success = book->saveRaw(&data, &size);
	if (success) {
		if (!data) {
			pThreadData->mExtraErrorText = "No data read from file";
			return METHOD_FAILED;
		}
		
		// Allocate memory with Omnis
		omSize = static_cast<qlong>(size);
		omData = static_cast<qbyte*>(MEMmalloc(omSize));
		
		// Copy data from LibXL into Omnis memory
		MEMmovel( data, omData,omSize);
		
		binVal.setBinary(fftBinary, omData, omSize);
		ECOsetParameterChanged( pThreadData->mEci, 1 ); // Notify Omnis that the binary data in the parameter changed
	}
	
	getEXTFldValFromBool(retVal, success);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Add a new sheet object
OmnisTools::tResult NVObjWorkbook::methodAddSheet( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	EXTfldval nameVal, objVal;
	
	// Read parameters.  
	// Param 1: Name of Sheet
	if ( getParamVar(pThreadData, 1, nameVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, sheetName, is required.";
		return ERR_BAD_PARAMS;
	} 
	
	// (Optional) Param 2: Sheet object to copy
	Sheet *sheet = 0, *retSheet = 0;
	NVObjWorksheet *omSheet;
	if ( getParamVar(pThreadData, 2, objVal) == qtrue ) {
		// Optional: Get Sheet object for use when creating the new value
		omSheet = getObjForEXTfldval<NVObjWorksheet>(pThreadData, objVal);
		sheet = omSheet->getSheet();
	}
	
	// Call LibXL to add sheet object
	if (sheet) {
		retSheet = book->addSheet(getWStringFromEXTFldVal(nameVal).c_str(), sheet);
	} else {
		retSheet = book->addSheet(getWStringFromEXTFldVal(nameVal).c_str());
	}
	if (!retSheet) {
		pThreadData->mExtraErrorText = str(format("Worksheet object returned empty. Error: %s") % book->errorMessage());
		return ERR_METHOD_FAILED;
	}
	
	// Create omnis representation of sheet and add neccesary pointers
	NVObjWorksheet* retOmnisSheet = createNVObj<NVObjWorksheet>(pThreadData);
	retOmnisSheet->setDependencies(numFormats, book, retSheet);
	retOmnisSheet->setErrorInstance(this->getInstance());
	
	EXTfldval retVal;
	getEXTFldValForObj<NVObjWorksheet>(retVal, retOmnisSheet);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Retrieve a sheet object
OmnisTools::tResult NVObjWorkbook::methodGetSheet( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	// Read parameters.  
	// Param 1: Index of sheet
	qlong sheetIndex;
	if ( getParamLong(pThreadData, 1, sheetIndex) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, sheetIndex, is required.";
		return ERR_BAD_PARAMS;
	} 
	if (sheetIndex < 1 || sheetIndex > (book->sheetCount()+1)) {
		pThreadData->mExtraErrorText = str(format("First parameter, sheetIndex, is out of range.  Must be between 1 and the number of Worksheets(%d)") % book->sheetCount());
		return ERR_BAD_PARAMS;
	}
	
	int index = static_cast<int>(sheetIndex) - 1; // Indexes in Omnis are always off of 1, but LibXL is off of 0.
	Sheet* sheet = book->getSheet(index);
	if (!sheet) {
		pThreadData->mExtraErrorText = str(format("Worksheet object returned empty. Error: %s") % book->errorMessage());
		return ERR_METHOD_FAILED;
	}
	
	// Create omnis representation of sheet and add neccesary pointers
	NVObjWorksheet* retOmnisSheet = createNVObj<NVObjWorksheet>(pThreadData);
	retOmnisSheet->setDependencies(numFormats, book, sheet);
	retOmnisSheet->setErrorInstance(this->getInstance());
	
	EXTfldval retVal;
	getEXTFldValForObj<NVObjWorksheet>(retVal, retOmnisSheet);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Delete a sheet
OmnisTools::tResult NVObjWorkbook::methodDelSheet( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	// Read parameters.  
	// Param 1: Index of sheet
	qlong sheetIndex;
	if ( getParamLong(pThreadData, 1, sheetIndex) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, sheetIndex, is required.";
		return ERR_BAD_PARAMS;
	} 
	if (sheetIndex < 1 || sheetIndex > (book->sheetCount()+1)) {
		pThreadData->mExtraErrorText = str(format("First parameter, sheetIndex, is out of range.  Must be between 1 and the number of Worksheets(%d)") % book->sheetCount());
		return ERR_BAD_PARAMS;
	}
	
	int index = static_cast<int>(sheetIndex) - 1; // Indexes in Omnis are always off of 1, but LibXL is off of 0.
	bool success = book->delSheet(index);
	
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, success);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Add a new format and return the format object
OmnisTools::tResult NVObjWorkbook::methodAddFormat( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	EXTfldval objVal;
	
	// (Optional) Param 1: Format object to copy
	Format* fmt = 0;
	Format* retFmt = 0;
	NVObjFormat *omFormat;
	if ( getParamVar(pThreadData, 1, objVal) == qtrue ) {
		// Optional: Get Format object for use when creating the new value
		omFormat = getObjForEXTfldval<NVObjFormat>(pThreadData, objVal);
		if (!omFormat) {
			pThreadData->mExtraErrorText = "Unrecognized first parameter, initFormat.";
			return ERR_METHOD_FAILED;
		}
		fmt = omFormat->getFormat();
	}
	
	// Call LibXL to add format object
	if (fmt) {
		retFmt = book->addFormat(fmt);
	} else {
		retFmt = book->addFormat();
	}
	if (!retFmt) {
		pThreadData->mExtraErrorText = str(format("Format object returned empty. Error: %s") % book->errorMessage());
		return ERR_METHOD_FAILED;
	}
	
	// Create omnis representation of format and add neccesary pointers
	NVObjFormat* retOmnisFormat = createNVObj<NVObjFormat>(pThreadData);
	retOmnisFormat->setDependencies(book, retFmt);
	retOmnisFormat->setErrorInstance(this->getInstance());
	
	EXTfldval retVal;
	getEXTFldValForObj<NVObjFormat>(retVal, retOmnisFormat);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Add a new font and return the font object
OmnisTools::tResult NVObjWorkbook::methodAddFont( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	EXTfldval objVal;
	
	// (Optional) Param 1: Font object to copy
	Font *font = 0, *retFont = 0;
	NVObjFont *omFont;
	if ( getParamVar(pThreadData, 1, objVal) == qtrue ) {
		// Optional: Get Font object for use when creating the new value
		omFont = getObjForEXTfldval<NVObjFont>(pThreadData, objVal);
		if (!omFont) {
			pThreadData->mExtraErrorText = "Unrecognized first parameter, initFont.";
			return ERR_METHOD_FAILED;
		}
		font = omFont->getFont();
	}
	
	// Call LibXL to add font object
	if (font) {
		retFont = book->addFont(font);
	} else {
		retFont = book->addFont();
	}
	if (!retFont) {
		pThreadData->mExtraErrorText = str(format("Font object returned empty. Error: %s") % book->errorMessage());
		return ERR_METHOD_FAILED;
	}
	
	// Create omnis representation of font and add neccesary pointers
	NVObjFont* retOmnisFont = createNVObj<NVObjFont>(pThreadData);
	retOmnisFont->setDependencies(book, retFont);
	retOmnisFont->setErrorInstance(this->getInstance());
	
	EXTfldval retVal;
	getEXTFldValForObj<NVObjFont>(retVal, retOmnisFont);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Add a new custom number format index for a format string
OmnisTools::tResult NVObjWorkbook::methodAddCustomNumFormat( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	// Read parameters
	// Param 1: Format to be used.  See: http://libxl.com/custom-format.html
	EXTfldval formatVal;
	if ( getParamVar(pThreadData, 1, formatVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, customNumFormat, is required.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL to add custom number format
	std::wstring formatString = getWStringFromEXTFldVal(formatVal);
	int formatNum = book->addCustomNumFormat(formatString.c_str());
	if (formatNum == 0) {
		pThreadData->mExtraErrorText = str(format("Unable to create custom number format.  Error: %s") % book->errorMessage() );
		return ERR_METHOD_FAILED;
	}
	
	// Return format number to caller
	EXTfldval retVal;
	getEXTFldValFromInt(retVal, formatNum);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Return a format string for a given custom number format index
OmnisTools::tResult NVObjWorkbook::methodCustomNumFormat( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	// Read parameters
	// Param 1: Format number that was created with $addCustomNumFormat
	qlong formatNum;
	if ( getParamLong(pThreadData, 1, formatNum) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, fmt, is required.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL to add custom number format
	int formatInt = static_cast<int>(formatNum);
	
	// Return format string to caller
	EXTfldval retVal;
	getEXTFldValFromWChar(retVal, book->customNumFormat(formatInt) );
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Return a format object for a given format index
OmnisTools::tResult NVObjWorkbook::methodFormat( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	// Read parameters
	// Param 1: Index of format
	qlong formatIndex;
	if ( getParamLong(pThreadData, 1, formatIndex) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, index, is required.";
		return ERR_BAD_PARAMS;
	}
	if (formatIndex < 1 || formatIndex > (book->formatSize()+1)) {
		pThreadData->mExtraErrorText = str(format("First parameter, index, is out of range.  Must be between 1 and size of Formats(%d)") % book->formatSize());
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL to retrieve format object
	int intIndex = static_cast<int>(formatIndex);
	Format* fmt = book->format(intIndex);
	if (!fmt) {
		pThreadData->mExtraErrorText = str(format("Format object returned empty. Error: %s") % book->errorMessage());
		return ERR_METHOD_FAILED;
	}
	
	// Create omnis representation of format and add neccesary pointers
	NVObjFormat* retOmnisFormat = createNVObj<NVObjFormat>(pThreadData);
	retOmnisFormat->setDependencies(book, fmt);
    retOmnisFormat->setErrorInstance(this->getInstance());
	
	// Return format object
	EXTfldval retVal;
	getEXTFldValForObj<NVObjFormat>(retVal, retOmnisFormat);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Return a font object for a given font index
OmnisTools::tResult NVObjWorkbook::methodFont( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	// Read parameters
	// Param 1: Index of format
	qlong fontIndex;
	if ( getParamLong(pThreadData, 1, fontIndex) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, index, is required.";
		return ERR_BAD_PARAMS;
	}
	if (fontIndex < 1 || fontIndex > (book->fontSize()+1)) {
		pThreadData->mExtraErrorText = str(format("First parameter, index, is out of range.  Must be between 1 and size of Fonts(%d)") % book->fontSize());
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL to retrieve format object
	int intIndex = static_cast<int>(fontIndex);
	Font* font = book->font(intIndex);
	if (!font) {
		pThreadData->mExtraErrorText = str(format("Font object returned empty. Error: %s") % book->errorMessage());
		return ERR_METHOD_FAILED;
	}
	
	// Create omnis representation of font and add neccesary pointers
	NVObjFont* retOmnisFont = createNVObj<NVObjFont>(pThreadData);
	retOmnisFont->setDependencies(book, font);
    retOmnisFont->setErrorInstance(this->getInstance());
	
	// Return font object
	EXTfldval retVal;
	getEXTFldValForObj<NVObjFont>(retVal, retOmnisFont);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

OmnisTools::tResult NVObjWorkbook::methodAddPicture( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	EXTfldval pathVal;
	
	// Param 1: Path to picture file (Support POSIX and HFS on Mac)
	if ( getParamVar(pThreadData, 1, pathVal) != qtrue )
		return ERR_BAD_PARAMS;
	if ( ensurePosixPath(pathVal) != qtrue ) {
		pThreadData->mExtraErrorText = str(format("Unable to convert HFS Path to Posix: \"%s\"") % getStringFromEXTFldVal(pathVal));
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL with file name
	std::wstring path = getWStringFromEXTFldVal(pathVal);
	int pictNum = book->addPicture(path.c_str());
	
	// Return format number to caller
	EXTfldval retVal;
	getEXTFldValFromInt(retVal, pictNum);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	if (pictNum == -1) {
		pThreadData->mExtraErrorText = str(format("Unable to add picture: \"%s\".  Error: %s") % getStringFromEXTFldVal(pathVal) % book->errorMessage() );
		return ERR_METHOD_FAILED;
	} else {
		return METHOD_DONE_RETURN;
	}
}

OmnisTools::tResult NVObjWorkbook::methodAddPictureRaw( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	EXTfldval binVal;
	
	// Param 1: Binary with picture data
	if ( getParamVar(pThreadData, 1, binVal) != qtrue )
		return ERR_BAD_PARAMS;
	
	qlong omSize = binVal.getBinLen();
	if (omSize == 0) {
		pThreadData->mExtraErrorText = "No binary data in parameter";
		return ERR_BAD_PARAMS;
	}
	
	//  Interpret omnis binary into LibXL binary
	qbyte* omData = static_cast<qbyte*>(MEMmalloc(omSize));
	
	binVal.getBinary(omSize, omData, omSize);
	
	unsigned size = static_cast<unsigned>(omSize);
	const char* data = reinterpret_cast<const char*> (omData);
	
	// Call LibXL with binary data
	int pictNum = book->addPicture2(data, size);
	
	// Return format number to caller
	EXTfldval retVal;
	getEXTFldValFromInt(retVal, pictNum);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	if (pictNum == -1) {
		pThreadData->mExtraErrorText = str(format("Unable to add picture from binary.  Error: %s") % book->errorMessage() );
		return ERR_METHOD_FAILED;
	} else {
		return METHOD_DONE_RETURN;
	}
}

// Method for $populateFromList.  Create a simple spreadsheet from an Omnis list.
OmnisTools::tResult NVObjWorkbook::methodPopulateFromList( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	EXTqlist listVal, headerVal, formatVal;
	qshort styleVal = POP_STYLE_NO_HEADER;
	
	if( getParamList(pThreadData, 1, listVal, qfalse) != qtrue ) { 
		pThreadData->mExtraErrorText = "List in 1st parameter not recognized.";
		return ERR_BAD_PARAMS;
	}
	
	getParamShort(pThreadData, 2, styleVal); // Optional
	if (styleVal == POP_STYLE_CUSTOM_HEADER || styleVal == POP_STYLE_CUSTOM_HEADER_FMT) {
		if( getParamList(pThreadData, 3, headerVal, qtrue) != qtrue ) {
			pThreadData->mExtraErrorText = "Missing header row in 3rd parameter.";
			return ERR_BAD_PARAMS;
		}
	}
    
    if (styleVal == POP_STYLE_CUSTOM_HEADER_FMT) {
		if( getParamList(pThreadData, 4, formatVal, qtrue) != qtrue ) {
			pThreadData->mExtraErrorText = "Missing format row in 4th parameter.";
			return ERR_BAD_PARAMS;
		}
	}
	
	if (styleVal < POP_STYLE_NO_HEADER || styleVal > POP_STYLE_CUSTOM_HEADER_FMT ) {
		pThreadData->mExtraErrorText = str(format("Style %d is not valid.") % styleVal);
		return ERR_NOT_SUPPORTED;
	}
	
    // Remove all previous sheets
    for(int x = 0; x < book->sheetCount(); ++x)
        book->delSheet(x);
    
	Sheet* sheet = book->addSheet(L"Sheet1");
    if (!sheet) {
        pThreadData->mExtraErrorText = "Unable to get worksheet for workbook.";
		return METHOD_FAILED;
    }
    
	// Loop list and add all data
	qlong rowCount = listVal.rowCnt();
	qlong rowOffset = 0;
	qshort colCount = listVal.colCnt();
	EXTfldval colVal;
	str255 colName;
	const qshort charFormat = dpFcharacter;
	Font* headerFont = 0;
	Format* headerFmt = 0;
	
	// Setup header information (if applicable)
	if (styleVal == POP_STYLE_DEFINITION_HEADER || styleVal == POP_STYLE_CUSTOM_HEADER || styleVal == POP_STYLE_CUSTOM_HEADER_FMT) {
		headerFont = book->addFont();
		headerFont->setBold();
		headerFont->setSize(11);
		headerFmt = book->addFormat();
        headerFmt->setWrap(true);
		headerFmt->setFont(headerFont);
		headerFmt->setBorderBottom(BORDERSTYLE_MEDIUM);
		
		rowOffset += 1; // Set offset for data
	}
	
	// Write header information
	int headerRow = 1;
	if (styleVal == POP_STYLE_DEFINITION_HEADER) {
		// Use column names
		for ( qshort curCol = 1 ; curCol <= colCount ; curCol++ ) {
			listVal.getCol(curCol, qfalse, colName);
			colVal.setChar(colName, charFormat);
			writeValueToSheet(sheet, headerRow, curCol, colVal, headerFmt);
		}
	} else if (styleVal == POP_STYLE_CUSTOM_HEADER || styleVal == POP_STYLE_CUSTOM_HEADER_FMT) {
		for ( qshort curCol = 1 ; curCol <= colCount ; curCol++ ) {
			headerVal.getColVal( 1, curCol, colVal );
			
			writeValueToSheet(sheet, headerRow, curCol, colVal, headerFmt);
		}
	}
    
    // Setup the formats needed for each column
    Format* curFormat;
    NVObjFormat* omFormat;
    std::map<qlong, Format*> formatLookup;
    std::map<qlong, Format*>::iterator fmtIterator;
    
    // For the format compatible style build the list of formats
    if (styleVal == POP_STYLE_CUSTOM_HEADER_FMT) {
        for ( qshort curCol = 1 ; curCol <= colCount ; curCol++ ) {
			formatVal.getColVal( 1, curCol, colVal );
            if (getType(colVal).valType == fftObject || getType(colVal).valType == fftObjref) {
                omFormat = getObjForEXTfldval<NVObjFormat>(pThreadData, colVal);
                if( omFormat ) {
                    curFormat = omFormat->getFormat();
                    if (curFormat) {
                        formatLookup[curCol] = curFormat;
                    } else {
                        pThreadData->mExtraErrorText = str(format("Found object in 4th parameter, formatRow, but couldn't resolve object contents in column %d.") % curCol);
                        return ERR_BAD_PARAMS; 
                    }
                }
            } else {
                pThreadData->mExtraErrorText = "Invalid definition in contents of 4th parameter, formatRow.  Expected kObject.";
                return ERR_BAD_PARAMS; 
            }
			
		}
    }
    
	// Write data
    for ( qshort curCol = 1 ; curCol <= colCount ; curCol++ ) {
        // Lookup format
        fmtIterator = formatLookup.find(curCol);
        if (fmtIterator != formatLookup.end()) {
            curFormat = fmtIterator->second;
        } else {
            curFormat = 0;
        }
        
        // Write entire column
        for ( qlong curRow = 1; curRow <= rowCount; ++curRow ) {
			listVal.getColVal( curRow, curCol, colVal );
			writeValueToSheet(sheet, rowOffset+curRow, curCol, colVal, curFormat);
		}
	}
	
	return METHOD_DONE_RETURN;
}

// Helper function for methodExtractList to determine the definition required for a value
void addColForValue(EXTqlist* listVal, EXTfldval& fVal, str255* colName) {
	
	FieldValType fValType = getType(fVal);
	
	if (fValType.valType == fftCharacter) {
		listVal->addCol(fValType.valType, fValType.valSubType, 10000000, colName);
	} else {	
		listVal->addCol(fValType.valType, fValType.valSubType, 0, colName);
	}
}

// Method for $extractList - Gets an Omnis list from an Excel file
OmnisTools::tResult NVObjWorkbook::methodExtractList( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	
	// Determine style
	qshort styleVal = EXTRACT_STYLE_NO_HEADER;
	getParamShort(pThreadData, 1, styleVal); // Optional
	if (styleVal < EXTRACT_STYLE_NO_HEADER || styleVal > EXTRACT_STYLE_MAP_FROM_HEADER)
		return ERR_BAD_PARAMS;
	
	// Optional: Read list to use for definition of new list
	EXTqlist defineList;
	if (styleVal == EXTRACT_STYLE_CUSTOM_DEFINE || styleVal == EXTRACT_STYLE_CUSTOM_DEFINE_IH || styleVal == EXTRACT_STYLE_MAP_FROM_HEADER) {
		if( getParamList(pThreadData, 2, defineList, qfalse) != qtrue ) { 
			pThreadData->mExtraErrorText = "List in 2nd parameter is not recognized.";
			return ERR_BAD_PARAMS;
		}
	}
	
	if ( book->sheetCount() < 1) {
		pThreadData->mExtraErrorText = str(format("Invalid sheet count.  Expected greater then 0, got %d.") % book->sheetCount() );
		return ERR_BAD_CALCULATION;
	}
	
	// Sheet layout
	Sheet* sheet = book->getSheet(0);
	
	int rowStart = sheet->firstRow(); 
	int rowCount = (sheet->lastRow()+1) - rowStart;
	
	const int MAX_COLS = 255;
	int colStart = sheet->firstCol(); 
	int colCount = sheet->lastCol() - colStart;
	colCount = colCount > MAX_COLS ? MAX_COLS : colCount;
	
	// Omnis list and list column names
	EXTqlist* listVal = new EXTqlist(listVlen);
	EXTfldval colVal, defVal, retVal;
	str255 colName;
	qlong rowNum;
	
	// Alter data position if using or ignoring a header
	qlong headerRow = rowStart + 1; // Index header row off of 1
	if (styleVal != EXTRACT_STYLE_NO_HEADER && styleVal != EXTRACT_STYLE_CUSTOM_DEFINE) {
		rowStart += 1;
		rowCount -= 1;
	}
	
	// Column Map used for determining which column in the data maps to which column in the list
	std::map<qshort, qshort> colMap;
	std::map<std::string, qshort> defMap;
	std::string colKey;
	qshort defCols;
	
	// For types
	ffttype basetype;
	qshort subtype;
	qlong sublen;
	
	// Define list
	
	if (styleVal == EXTRACT_STYLE_MAP_FROM_HEADER) {
		// When mapping the data to a specific definition the list must be copied
		// and then the definition mapped in the loop below
		
		defCols = defineList.colCnt(); 
		for (qshort mapCol = 1; mapCol <= defCols && mapCol <= MAX_COLS; ++mapCol) {
			
			defineList.getCol(mapCol, qfalse, colName);
			colVal.setChar(colName, dpDefault);
			
			// Add column mapping based on column name
			colKey = getStringFromEXTFldVal(colVal);
			defMap[colKey] = mapCol;
			
			colVal.getChar(colName);
			defineList.getColType( mapCol, basetype, subtype, sublen );
			listVal->addCol(basetype, subtype, sublen, &colName);
		}
	}
	
	// Loop all columns and define list with first row of data in the spreadsheet
	int colPos;
	for ( qshort curCol = 1; curCol <= colCount ; ++curCol ) {
		colVal.setEmpty(fftCharacter, dpFcharacter);
		colName = qnil; //Zero out colName
		colPos = colStart+curCol;
		
		if (styleVal == EXTRACT_STYLE_USE_HEADER || styleVal == EXTRACT_STYLE_MAP_FROM_HEADER) {
			// Read header from list
			readValueFromSheet(colVal, sheet, headerRow, colPos);
		} else if (styleVal == EXTRACT_STYLE_CUSTOM_DEFINE || styleVal == EXTRACT_STYLE_CUSTOM_DEFINE_IH) {
			// Get column name from passed in list
			defineList.getCol(curCol, qfalse, colName);
		} 
		
		if (!colName) {
			// If there is no value in the header or the default header option is selected then use Excel column name
			if (colVal.isEmpty()) {
				getColNameForColNumber(colVal, colPos);
			}
			colVal.getChar(colName, qtrue);
		}
		
		if (styleVal == EXTRACT_STYLE_MAP_FROM_HEADER) {
			// List is already defined just need to map columns together
			colKey = getStringFromEXTFldVal(colVal);
			colMap[colPos] = defMap[colKey];
		} else {
			colMap[colPos] = colPos; // 1 -> 1 Column Mapping
			
			readValueFromSheet(defVal, sheet, rowStart+1, colPos); // Read first row of data to determine type
			addColForValue(listVal, defVal, &colName);
		}
	}

	// Read data into list
	for( qlong curRow = 1; curRow < rowCount; ++curRow ) {
		rowNum = listVal->insertRow(0); // Insert after last row
		if (!rowNum)
			break;
			
		// Valid row -- add column data
		for ( qshort curCol = 1 ; curCol <= colCount ; ++curCol ) {
			listVal->getColValRef( curRow, colMap[curCol], colVal, qtrue );
			readValueFromSheet(colVal, sheet, rowStart+curRow, colStart+curCol);
		}
	}
	
	// Setup return value
	retVal.setList(listVal, qtrue, qfalse); 
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

/**************************************************************************************************
 **                              INTERNAL      METHODS                                           **
 **************************************************************************************************/

// Book requires that the release method be called before it's deleted.  
void BookDeleter( Book* ptr) {
	if (ptr) {
		ptr->release();
	}
}

// Internal initialize method that sets the proper book type 
void NVObjWorkbook::initialize() {
	
	// Get appropriate Book class
	if (fileType == XLSX_FORMAT) {
		book = shared_ptr<Book>( xlCreateXMLBook(), BookDeleter );
	} else {
		book = shared_ptr<Book>( xlCreateBook(), BookDeleter );
	}
	
	// Setup licensing
	book->setKey(licenseName.c_str(), licenseKey.c_str());
	
	// Construct number format manager
	numFormats = shared_ptr<NumFormatMgr>(new NumFormatMgr(book));
	
	return;
}

// Get double value used as for a date
XLDateFormat NVObjWorkbook::getDateForEXTFldVal(EXTfldval& fVal, boost::shared_ptr<Book> bk) {
	FieldValType fValType = getType(fVal);
	
	XLDateFormat retDate;
	datestamptype convDate;
	qshort def = fValType.valSubType; // Used for subformat of date
	fVal.getDate(convDate, def);
	
	int year = 0, month = 0, day = 0, hour = 0, min = 0, sec = 0, msec = 0;
	bool setDate = false, setTime = false, setSec = false;
	
	// Determine what kind of date to assign based on the type of the field
	
	if (convDate.mDateOk && fValType.valSubType != dpFtime) {
		year  = static_cast<int>(convDate.mYear);
		month = static_cast<int>(convDate.mMonth);
		day   = static_cast<int>(convDate.mDay);
		setDate = true;
	}
	if (convDate.mTimeOk && fValType.valSubType != dpFdate1900 && fValType.valSubType != dpFdate1980 && fValType.valSubType != dpFdate2000) {
		hour = static_cast<int>(convDate.mHour);
		min = static_cast<int>(convDate.mMin);
		setTime = true;
	}
	if (convDate.mSecOk && fValType.valSubType != dpFdate1900 && fValType.valSubType != dpFdate1980 && fValType.valSubType != dpFdate2000) {
		sec = static_cast<int>(convDate.mSec);
		setSec = true;
	}
	if (convDate.mHunOk && fValType.valSubType != dpFdate1900 && fValType.valSubType != dpFdate1980 && fValType.valSubType != dpFdate2000) {
		msec = static_cast<int>(convDate.mHun);
		msec = msec * 10; // Convert from hundreths (Omnis) to milliseconds (LibXL)
	}
	
	// Get double value
	retDate.date = bk->datePack(year, month, day, hour, min, sec, msec);
	
	// Determine format
	if (setDate && setTime) {
		retDate.format = NUMFORMAT_CUSTOM_MDYYYY_HMM;
	} else if (setDate) {
		retDate.format = NUMFORMAT_DATE;
	} else if (setTime) {
		if (setSec) {
			retDate.format = NUMFORMAT_CUSTOM_HMM_AM;
		} else {
			retDate.format = NUMFORMAT_CUSTOM_HMMSS_AM;
		}
	}
	
	return retDate;
}

// Get date from an excel double value (packed date)
void NVObjWorkbook::getEXTFldValForDate(EXTfldval& fVal, XLDateFormat packedDate, boost::shared_ptr<Book> bk) {
	datestamptype convDate;
	qshort dp = dpFdtimedefault; // Used for subformat of date
	
	int year = 0, month = 0, day = 0, hour = 0, min = 0, sec = 0, msec = 0;
	
	bk->dateUnpack(packedDate.date, &year, &month, &day, &hour, &min, &sec, &msec);
	
	if (year != 0 || month != 0 || day != 0) {
		convDate.mYear  = static_cast<qshort>(year);
		convDate.mMonth = static_cast<char>(month);
		convDate.mDay = static_cast<char>(day);
		convDate.mDateOk = static_cast<char>(qtrue);
		
		dp = dpFdate1980;  // Set as date only before performing time checks
	}
	
	if (hour != 0 || min != 0 || sec != 0 || msec != 0 ) {
		convDate.mHour = static_cast<char>(hour);
		convDate.mMin = static_cast<char>(min);
		convDate.mTimeOk = static_cast<char>(qtrue);
		
		convDate.mSec = static_cast<char>(sec);
		convDate.mSecOk = static_cast<char>(qtrue);
		
		convDate.mHun = static_cast<char>(msec/10);
		convDate.mHunOk = static_cast<char>(qtrue);
		
		if (dp == dpFdate1980) {
			dp = dpFdtimedefault; // Have a date, switch to date/time
		} else {
			dp = dpFtime;
		}
	}
	
	fVal.setDate(convDate, dp);
}

// Read a specified sheet, row and column into an Omnis EXTfldval
bool NVObjWorkbook::readValueFromSheet(EXTfldval& cellValue, libxl::Sheet* sheet, qlong curRow, qshort curCol) {
	int row = static_cast<int>(curRow) - 1; // LibXL is indexed off 0
	int col = static_cast<int>(curCol) - 1; // LibXL is indexed off 0
	
	XLDateFormat d;
	Format* fmt;
	double readDouble, testDouble;
	long readLong, testLong;
	
	if ( sheet->isDate(row, col) ) {
		// Date
		d.date = sheet->readNum(row, col, &fmt);
		if (fmt) {
			d.format = fmt->numFormat();
		}
		
		getEXTFldValForDate(cellValue, d, book);
	} else {
		// Normal type
		switch ( sheet->cellType(row,col) ) {
			case CELLTYPE_EMPTY:
			case CELLTYPE_BLANK:
			case CELLTYPE_ERROR:
				getEXTFldValFromChar(cellValue, "");
				break;
			case CELLTYPE_STRING:
				getEXTFldValFromWChar(cellValue, sheet->readStr(row,col));
				break;
			case CELLTYPE_BOOLEAN:
				getEXTFldValFromBool(cellValue, sheet->readBool(row, col));
				break;
			case CELLTYPE_NUMBER:
				readDouble = sheet->readNum(row,col);
				testLong = static_cast<long>(readDouble);
				
				// There is no way to determine what the original data type was other then to see if there are decimal places.
				testDouble = readDouble - testLong; // If there is a remainder then it's a double
				if (testDouble == 0.0) {
					readLong = static_cast<long>(readDouble);
					getEXTFldValFromLong(cellValue, readLong);
				} else {
					getEXTFldValFromDouble(cellValue, readDouble);
				}
				
				break;
			default:
				break;
		}
	}
	
	return true;
}

// Write any Omnis EXTfldval value to a specified sheet, row and column.
bool NVObjWorkbook::writeValueToSheet(Sheet* sheet, qlong curRow, qshort curCol, EXTfldval& cellValue, Format* pFormat) {
	int row = static_cast<int>(curRow) - 1; // LibXL is indexed off 0
	int col = static_cast<int>(curCol) - 1; // LibXL is indexed off 0
	
	Format *dateFmt = 0;
	XLDateFormat d;

	// Write value based on EXTfldval type
	FieldValType cellType = getType(cellValue);
	switch (cellType.valType) {
		case fftCharacter:
			sheet->writeStr(row, col, getWStringFromEXTFldVal(cellValue).c_str(), pFormat);
			break;
		case fftInteger:
			sheet->writeNum(row, col, getLongFromEXTFldVal(cellValue), pFormat);
			break;
		case fftNumber:
			sheet->writeNum(row, col, getDoubleFromEXTFldVal(cellValue), pFormat);
			break;
		case fftBoolean:
			sheet->writeBool(row, col, getBoolFromEXTFldVal(cellValue), pFormat);
			break;
		case fftDate:
			 d = getDateForEXTFldVal(cellValue, book);
			
			// Determine and/or reuse format
			if (pFormat) {
				dateFmt = pFormat;
			} else {
				dateFmt = numFormats->getFormat(d.format);
			}
			
			// Only write data if data exists
			if (d.date != 0.0) {
				sheet->writeNum(row, col, d.date, dateFmt);
			} else {
				sheet->writeBlank(row, col, dateFmt);
			}
			
			break;
		default:
			break;
	}

	return true;
}

// Conversion array for getting Excel column name from a column number
const static int letterConvert[] = {L'A',L'B',L'C',L'D',L'E',L'F',L'G',L'H',L'I',L'J',L'K',L'L',L'M',
	                                L'N',L'O',L'P',L'Q',L'R',L'S',L'T',L'U',L'V',L'X',L'Y',L'Z'};

// Get a column number from an EXTfldval -- allows for easy conversion from parameter to integer for LibXL
int NVObjWorkbook::getColNumber(EXTfldval& colVal) {
	int col;
	
	// Based on type set the column number
	if (getType(colVal).valType == fftCharacter) {
		col = NVObjWorkbook::getColNumberForColName(colVal);
	} else if (getType(colVal).valType == fftInteger) {
		col = getIntFromEXTFldVal(colVal);
	} else {
		col = -1;
	}
	
	return col;
}

// Get Excel 'A', 'B', 'C' column name for column number
void NVObjWorkbook::getColNameForColNumber(EXTfldval& colName, int colNum) {	
	std::wstring retString;
	if (colNum > 26) {
		// If the column number is greater then 26 then the first 
		// letter is a multiple of 26.
		int last = colNum % 26;
		colNum = colNum/26;
		
		retString = letterConvert[colNum-1]; // First letter
		retString += letterConvert[last-1];  // Second letter
	} else if (colNum > 0) {
		retString = letterConvert[colNum-1]; // Single letter
	}

	getEXTFldValFromWString(colName, retString);
	
	return;
}

// Get column number for Excel column name, 'A', 'B', 'C', etc..
int NVObjWorkbook::getColNumberForColName(EXTfldval& colName) {
	
	std::wstring colString = getWStringFromEXTFldVal(colName);
	int colNumber = 0;
	
	if (colString.length() == 1) {
		for(int i = 0; i <= 25; ++i) {
			if (colString[0] == letterConvert[i]) {
				colNumber = i+1; // Index off 1 for Omnis' sake
				break;
			}
		}
	} else if (colString.length() == 2) {
		int first = -1, second = -1;
		
		// Loop and find the position of each letter
		for(int i = 0; i <= 25; ++i) {
			if (colString[0] == letterConvert[i]) {
				first = i+1;  // Index off 1 for Omnis' sake
			} else if (colString[0] == letterConvert[i]) {
				second = i+1;  // Index off 1 for Omnis' sake
			}
			if (first > -1 && second > -1)
				break; // Found both values, leave loop
		}
		colNumber = (first*26) + second;
	}
	
	return colNumber;
}

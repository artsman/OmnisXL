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
#include "CellObj.he"
#include "FormatObj.he"
#include "FontObj.he"

#include <boost/format.hpp>

using namespace OmnisTools;
using namespace libxl;
using boost::shared_ptr;
using boost::format;

static const qlong TYPE_CONST_START = 23100, 
                     TYPE_CONST_END = 23106;

static const qlong ERROR_CONST_START = 23107, 
				   ERROR_CONST_END   = 23114;


/**************************************************************************************************
 **                       CONSTRUCTORS / DESTRUCTORS                                             **
 **************************************************************************************************/

NVObjCell::NVObjCell(qobjinst qo, tThreadData *pThreadData) : NVObjBase(qo)
{ }

NVObjCell::~NVObjCell()
{ }

void NVObjCell::copy( NVObjCell* pObj ) {
	// Copy in super class (This does *this = *pObj)
	NVObjBase::copy(pObj);
	
	// Copy the contents of the object into the new object
	book = pObj->book;
	sheet = pObj->sheet;
    numFormats = pObj->numFormats;
    row = pObj->row;
    col = pObj->col;
}

/**************************************************************************************************
 **                              PROPERTY DECLERATION                                            **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project it is also used as the Unique ID.
const static qshort cCPropertyType      = 4100,
                    cCPropertyError     = 4101,
                    cCPropertyFormat    = 4102,
                    cCPropertyValue     = 4103,
                    cCPropertyFont      = 4104,
                    cCPropertyRow       = 4105,
                    cCPropertyCol       = 4106,
                    cCPropertyColLet    = 4107,
                    cCPropertyComment   = 4108,
                    cCPropertyIsFormula = 4109,
                    cCPropertyIsDate    = 4110;

/**************************************************************************************************
 **                               METHOD DECLERATION                                             **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project is also used as the Unique ID.
const static qshort cCMethodSetPicture     = 4000,
					cCMethodSetValue       = 4001,
                    cCMethodClear          = 4002,
					cCMethodSetComment     = 4003,
                    cCMethodReadFormula    = 4004,
                    cCMethodWriteFormula   = 4005;

/**************************************************************************************************
 **                                 INSTANCE METHODS                                             **
 **************************************************************************************************/

// Call a method
qlong NVObjCell::methodCall( tThreadData* pThreadData )
{
	qshort funcId = (qshort)ECOgetId(pThreadData->mEci);
	qshort paramCount = ECOgetParamCount(pThreadData->mEci);

	if (!checkDependencies()) {
		pThreadData->mExtraErrorText = "Internal Error.  Object dependencies are not setup.";
		callErrorMethod( pThreadData, ERR_METHOD_FAILED );
		return qfalse;
	}
	
	tResult result = ERR_OK;
	
	switch( funcId )
	{
		case cCMethodSetPicture:
			pThreadData->mCurMethodName = "$setPicture";
			result = methodSetPicture(pThreadData, paramCount);
			break;
		case cCMethodSetValue:
			pThreadData->mCurMethodName = "$setValue";
			result = methodSetValue(pThreadData, paramCount);
			break;
		case cCMethodClear:
			pThreadData->mCurMethodName = "$clear";
			result = methodClear(pThreadData, paramCount);
			break;
		case cCMethodSetComment:
			pThreadData->mCurMethodName = "$setComment";
			result = methodSetComment(pThreadData, paramCount);
			break;
        case cCMethodReadFormula:
            pThreadData->mCurMethodName = "$readFormula";
            result = methodReadFormula(pThreadData, paramCount);
            break;
        case cCMethodWriteFormula:
            pThreadData->mCurMethodName = "$writeFormula";
            result = methodWriteFormula(pThreadData, paramCount);
	}
	
	// Check for errors to notify Omnis developer in $error
	callErrorMethod( pThreadData, result );

	return 0L;
}

/**************************************************************************************************
 **                                PROPERTIES                                                    **
 **************************************************************************************************/

// Assignability of properties
qlong NVObjCell::canAssignProperty( tThreadData* pThreadData, qlong propID ) {
	switch (propID) {
		case cCPropertyType:
		case cCPropertyError:
		case cCPropertyRow:
		case cCPropertyCol:
		case cCPropertyColLet:
		case cCPropertyComment:
        case cCPropertyIsFormula:
        case cCPropertyIsDate:
			return qfalse;
		case cCPropertyFormat:
		case cCPropertyValue:
		case cCPropertyFont:
			return qtrue;
		default:
			return qfalse;
	}
}

// Method to retrieve a property of the object
qlong NVObjCell::getProperty( tThreadData* pThreadData ) 
{
	EXTfldval fValReturn;
	if (!checkDependencies()) {
		ECOaddParam(pThreadData->mEci, &fValReturn);
		return qtrue;
	}
	
	qlong constAssign = 0;
	NVObjFormat* omFmt = 0;
	NVObjFont* omFont = 0;
	Font* font;
	Format* fmt;
	
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cCPropertyType:
			if ( sheet->isDate(row, col) ) {
				constAssign = kCellTypeDate;
			} else {
				// Normal type
				switch ( sheet->cellType(row,col) ) {
					case CELLTYPE_EMPTY:
						constAssign = kCellTypeEmpty;
					case CELLTYPE_BLANK:
						constAssign = kCellTypeBlank;
					case CELLTYPE_ERROR:
						constAssign = kCellTypeError;
					case CELLTYPE_STRING:
						constAssign = kCellTypeString;
					case CELLTYPE_BOOLEAN:
						constAssign = kCellTypeBoolean;
					case CELLTYPE_NUMBER:
						constAssign = kCellTypeNumber;
					default:
						break;
				}
			}
			
			getEXTFldValFromConstant(fValReturn, TYPE_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cCPropertyError:
			constAssign = getOmnisErrorTypeConstant(sheet->readError(row, col));
			
			getEXTFldValFromConstant(fValReturn, ERROR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cCPropertyFormat:
			omFmt = createNVObj<NVObjFormat>(pThreadData);
			if (omFmt) {
				fmt = sheet->cellFormat(row,col);
				if (fmt) {
					omFmt->setDependencies(book, fmt);
					omFmt->setErrorInstance(mErrorInst);
				}
				getEXTFldValForObj<NVObjFormat>(fValReturn, omFmt);
			}
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cCPropertyValue:
			getEXTfldvalForCell(fValReturn);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cCPropertyFont:
			omFont = createNVObj<NVObjFont>(pThreadData);
			if (omFont && sheet) {
				fmt = sheet->cellFormat(row,col);
				if (fmt) {
					font = fmt->font();
					if (font) {
						omFont->setDependencies(book, font);
						omFont->setErrorInstance(mErrorInst);
						getEXTFldValForObj<NVObjFont>(fValReturn, omFont);
					}
				} else {
					// No format.  Return empty font object.
					getEXTFldValForObj<NVObjFont>(fValReturn, omFont);
				}

			}
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cCPropertyRow:
			getEXTFldValFromInt(fValReturn, row+1); // Show Omnis user indexes based off 1
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cCPropertyCol:
			getEXTFldValFromInt(fValReturn, col+1); // Show Omnis user indexes based off 1
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cCPropertyColLet:
			NVObjWorkbook::getColNameForColNumber(fValReturn, col);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
			
		case cCPropertyComment:
			getEXTFldValFromWChar(fValReturn, sheet->readComment(row, col) );
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
        case cCPropertyIsFormula:
            getEXTFldValFromBool(fValReturn, sheet->isFormula(row, col));
            ECOaddParam(pThreadData->mEci, &fValReturn);
            break;
        case cCPropertyIsDate:
            getEXTFldValFromBool(fValReturn, sheet->isDate(row, col));
            ECOaddParam(pThreadData->mEci, &fValReturn);
            break;
	}

	return 1L;
}

// Method to set a property of the object
qlong NVObjCell::setProperty( tThreadData* pThreadData )
{
	if (!checkDependencies()) {
		return qfalse;
	}
	
	// Retrieve value to set for property, always in first parameter
	EXTfldval fVal;
	if( getParamVar( pThreadData->mEci, 1, fVal) == qfalse ) 
		return qfalse;

	NVObjFormat* omFmt = 0;
	NVObjFont* omFont = 0;
	Format* fmt;
	Font* font;
	
	// Assign to the appropriate property
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cCPropertyFormat:
			omFmt = getObjForEXTfldval<NVObjFormat>(pThreadData, fVal);
			if (omFmt) {
				fmt = omFmt->getFormat();
				if (fmt) {
					sheet->setCellFormat(row, col, fmt);
				}
			}
			break;
			
		case cCPropertyValue:
			setCellForEXTfldval(fVal);
			break;
			
		case cCPropertyFont:
			omFont = getObjForEXTfldval<NVObjFont>(pThreadData, fVal);
			if (omFont) {
				font = omFont->getFont();
				if (font) {
					fmt = sheet->cellFormat(row,col);
					if (!fmt) {
                        fmt = book->addFormat();
                        sheet->setCellFormat(row, col, fmt);
                    }
                    if (fmt) {
                         fmt->setFont(font);
                    }
				}
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
ECOparam cCellMethodsParamsTable[] = 
{
	// $setPicture
	4900, fftInteger, 0, 0,
	4901, fftInteger, 0, 0,
	4902, fftInteger, 0, 0,
	// $setValue
	4903, fftInteger, 0, 0,
	4904, fftObject,  EXTD_FLAG_PARAMOPT, 0,
	4905, fftInteger, EXTD_FLAG_PARAMOPT, 0,
	// $setComment
	4906, fftInteger, 0, 0,
	4907, fftInteger, EXTD_FLAG_PARAMOPT, 0,
	4908, fftInteger, EXTD_FLAG_PARAMOPT, 0,
	4909, fftInteger, EXTD_FLAG_PARAMOPT, 0,
    // $writeFormula
    4910, fftCharacter, 0, 0
    
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
ECOmethodEvent cCellMethodsTable[] = 
{
	cCMethodSetPicture,   cCMethodSetPicture,   fftNone,      3,  &cCellMethodsParamsTable[0], 0, 0,
	cCMethodSetValue,     cCMethodSetValue,     fftBoolean,   3,  &cCellMethodsParamsTable[3], 0, 0,
	cCMethodClear,        cCMethodClear,        fftNone,      0,                            0, 0, 0,
	cCMethodSetComment,   cCMethodSetComment,   fftNone,      4,  &cCellMethodsParamsTable[6], 0, 0,
    cCMethodReadFormula,  cCMethodReadFormula,  fftCharacter, 0,                            0, 0, 0,
    cCMethodWriteFormula, cCMethodWriteFormula, fftBoolean,   1, &cCellMethodsParamsTable[10], 0, 0
};

// List of methods in Simple
qlong NVObjCell::returnMethods(tThreadData* pThreadData)
{
	const qshort cCMethodCount = sizeof(cCellMethodsTable) / sizeof(ECOmethodEvent);
	
	return ECOreturnMethods( gInstLib, pThreadData->mEci, &cCellMethodsTable[0], cCMethodCount );
}

/* PROPERTIES */

// Table of properties available from Simple
// Columns are:
// 1) Unique ID 
// 2) Name of Property (Resource #)
// 3) Return Type 
// 4) Flags describing the property
// 5) Additional Flags describing the property
// 6) Enum Start (Not sure what this does, 0 = disabled)
// 7) Enum Stop (Not sure what this does, 0 = disabled)
ECOproperty cCellPropertyTable[] = 
{
	cCPropertyType,      cCPropertyType,      fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, TYPE_CONST_START, TYPE_CONST_END,
	cCPropertyError,     cCPropertyError,     fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, ERROR_CONST_START, ERROR_CONST_END,
	cCPropertyFormat,    cCPropertyFormat,    fftObject,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cCPropertyValue,     cCPropertyValue,     fftInteger,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cCPropertyFont,      cCPropertyFont,      fftObject,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cCPropertyRow,       cCPropertyRow,       fftInteger,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cCPropertyCol,       cCPropertyCol,       fftInteger,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
    cCPropertyIsFormula, cCPropertyIsFormula, fftBoolean,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
    cCPropertyIsDate,    cCPropertyIsDate,    fftBoolean,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0
};

// List of properties in Simple
qlong NVObjCell::returnProperties( tThreadData* pThreadData )
{
	const qshort propertyCount = sizeof(cCellPropertyTable) / sizeof(ECOproperty);

	return ECOreturnProperties( gInstLib, pThreadData->mEci, &cCellPropertyTable[0], propertyCount );
}

/**************************************************************************************************
 **                              OMNIS VISIBLE METHODS                                           **
 **************************************************************************************************/

// Sets a picture with pictureId identifier at position row and col with scale factor. 
tResult NVObjCell::methodSetPicture( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: Picture ID
	qlong omPictID;
	if( getParamLong(pThreadData, 1, omPictID) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, row, is required.";
		return ERR_BAD_PARAMS;
	}
	int pictID = static_cast<int> (omPictID);
	
	EXTfldval scaleVal;
	double scale;
	qlong omHeight, omWidth;
	int height, width;
	
	if (pParamCount == 2) {
		// Using Sheet::setPicture2(int row, int col, int pictureID, double scale)
		
		// Param 2: Scale
		if( getParamVar(pThreadData, 2, scaleVal) != qtrue) {
			pThreadData->mExtraErrorText = "Second parameter, scale, is required.";
			return ERR_BAD_PARAMS;
		}
		scale = getDoubleFromEXTFldVal(scaleVal);
		
		// Call LibXL
		sheet->setPicture(row, col, pictID, scale);
		
	} else if (pParamCount == 3) {
		// Using Sheet::setPicture2(int row, int col, int pictureID, int width, int height)
		
		// Param 2: Height
		if( getParamLong(pThreadData, 2, omHeight) != qtrue) {
			pThreadData->mExtraErrorText = "Second parameter, height, is required.";
			return ERR_BAD_PARAMS;
		}
		height = static_cast<int> (omHeight);
		
		// Param 3: Width
		if( getParamLong(pThreadData, 3, omWidth) != qtrue) {
			pThreadData->mExtraErrorText = "Third parameter, width, is required.";
			return ERR_BAD_PARAMS;
		}
		width = static_cast<int> (omWidth);
		
		// Call LibXL
		sheet->setPicture(row, col, pictID, width, height);
	} else {
		pThreadData->mExtraErrorText = "Invalid number of parameters. Takes 2 or 3 parameters.";
		return ERR_BAD_PARAMS;
	}
	
	return METHOD_DONE_RETURN;
}

// Sets the value in the cell 
tResult NVObjCell::methodSetValue( tThreadData* pThreadData, qshort pParamCount ) {
	// Param 1: Scale
	EXTfldval setVal;
	if( getParamVar(pThreadData, 1, setVal) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, value, is required.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 2: (Optional) Format
	EXTfldval formatVal;
	NVObjFormat* omFmt;
	Format* fmt = 0;
	if( getParamVar(pThreadData, 2, formatVal) == qtrue) {
		omFmt = getObjForEXTfldval<NVObjFormat>(pThreadData, formatVal);
		if (!omFmt) {
			pThreadData->mExtraErrorText = "Unrecognized second parameter, format.  Expected Format object.";
			return ERR_BAD_PARAMS;
		}
		
		// Got a good object
		fmt = omFmt->getFormat();
	}
	
	// Param 3: (Optional) Type constant
	qlong omType;
	int typeConstant = -1;
	if( getParamLong(pThreadData, 3, omType) == qtrue) {
		typeConstant = static_cast<int>(omType);
	}
	
	// Set contents of Cell
	setCellForEXTfldval(setVal, typeConstant, fmt);

	return METHOD_DONE_RETURN;
}

// Clears out the cell 
tResult NVObjCell::methodClear( tThreadData* pThreadData, qshort pParamCount ) {
	sheet->clear(row, row, col, col);

	return METHOD_DONE_RETURN;
}

// Sets the comment for the cell
tResult NVObjCell::methodSetComment( tThreadData* pThreadData, qshort pParamCount ) {
	// Param 1: Comment
	EXTfldval commentVal;
	if( getParamVar(pThreadData, 1, commentVal) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, comment, is required.";
		return ERR_BAD_PARAMS;
	}
	std::wstring comment = getWStringFromEXTFldVal(commentVal);
	
	// Param 2: (Optional) Author
	EXTfldval authorVal;
	std::wstring author;
	bool passedAuthor = false;
	if( getParamVar(pThreadData, 2, authorVal) == qtrue) {
		author = getWStringFromEXTFldVal(authorVal);
		passedAuthor = true;
	}
	
	// Param 3: (Optional) Height
	qlong omHeight;
	int height = 75;
	if( getParamLong(pThreadData, 3, omHeight) == qtrue) {
		height = static_cast<int>(omHeight);
	}
	
	// Param 4: (Optional) Width
	qlong omWidth;
	int width = 129;
	if( getParamLong(pThreadData, 4, omWidth) == qtrue) {
		width = static_cast<int>(omWidth);
	}
	
	sheet->writeComment(row, col, comment.c_str(), passedAuthor ? author.c_str() : 0, height, width);

	return METHOD_DONE_RETURN;
}

// Reads a formula from the cell
tResult NVObjCell::methodReadFormula( tThreadData* pThreadData, qshort pParamCount ) {
    
    // Check if cell contains a formula
    if( sheet->isFormula(row,col) != true ) {
        pThreadData->mExtraErrorText = "Cannot return formula for the cell.  Cell does not contain a formula value.";
        return ERR_METHOD_FAILED;
    }
    
    // Return the formula
    EXTfldval retVal;
    getEXTFldValFromWChar(retVal, sheet->readFormula(row, col));
    ECOaddParam(pThreadData->mEci, &retVal);
    
    return METHOD_DONE_RETURN;
}

// Writes a formula to the cell
tResult NVObjCell::methodWriteFormula( tThreadData* pThreadData, qshort pParamCount ) {
    
    // Param 1: Formula
    EXTfldval formulaVal;
    if( getParamVar(pThreadData, 1, formulaVal) == qfalse ) {
        pThreadData->mExtraErrorText = "Parameter 1 is unrecognized.  Expected Character containing formula.";
        return ERR_METHOD_FAILED;
    }
    // Validate formula
    std::wstring formula = getWStringFromEXTFldVal(formulaVal);
    
    // Return status to caller
    EXTfldval retVal;
    bool success = sheet->writeFormula(row, col, formula.c_str());
    getEXTFldValFromBool(retVal, success);
    ECOaddParam(pThreadData->mEci, &retVal);
    
    return METHOD_OK;
}

/**************************************************************************************************
 **                              INTERNAL      METHODS                                           **
 **************************************************************************************************/

// Get and EXTfldval for the contents of the cell
void NVObjCell::getEXTfldvalForCell(EXTfldval& cellValue) {
	
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
		
		NVObjWorkbook::getEXTFldValForDate(cellValue, d, book);
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
}

// Set the contents of the cell to the value in an EXTfldval
void NVObjCell::setCellForEXTfldval(EXTfldval& cellValue, int cellConstant, Format* fmt) {
	
	std::wstring stringWrite;
	double doubleWrite;
	
	Format* dateFmt = 0;
	XLDateFormat d;
	
	// If no type is passed then type detection must occur automatically
	FieldValType cellType = getType(cellValue);
	if (cellConstant == -1) {
		switch (cellType.valType) {
			case fftCharacter:
				cellConstant = kCellTypeString;
				break;
			case fftInteger:
			case fftNumber:
				cellConstant = kCellTypeNumber;
				break;
			case fftBoolean:
				cellConstant = kCellTypeBoolean;
				break;
			case fftDate:
				cellConstant = kCellTypeDate;
			default:
				cellConstant = kCellTypeBlank;
				break;
		}
	}
	
	// Write values based on cell type
	switch (cellConstant) {
		case kCellTypeString:
			stringWrite = getWStringFromEXTFldVal(cellValue);
			sheet->writeStr(row, col, stringWrite.c_str(), fmt);
			break;
			
		case kCellTypeNumber:
			if (cellType.valType == fftInteger) {
				// LibXL uses double for all numbers.  Read from Omnis into a long first and then translate; it's more accurate.
				doubleWrite = static_cast<double>( getLongFromEXTFldVal(cellValue) );
			} else {
				doubleWrite = getDoubleFromEXTFldVal(cellValue);
			}
			
			sheet->writeNum(row, col, doubleWrite, fmt);
			break;
			
		case kCellTypeBoolean:
			sheet->writeBool(row, col, getBoolFromEXTFldVal(cellValue), fmt);
			break;
			
		case kCellTypeDate:
			d = NVObjWorkbook::getDateForEXTFldVal(cellValue, book);
			
			// Determine and/or reuse format
			if (fmt) {
				dateFmt = fmt;
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
		case kCellTypeBlank:
		case kCellTypeEmpty:
			sheet->writeBlank(row, col, fmt);
			break;
			
		default:
			break;
	}
	
	return;
}

bool NVObjCell::checkDependencies() {
	if (numFormats && book && sheet  && row>=0 && col>=0)
		return true;
	else
		return false;
}

shared_ptr<NumFormatMgr> NVObjCell::getNumFormatMgr() {
	return numFormats;
}

shared_ptr<Book> NVObjCell::getBook() {
	return book;
}

Sheet* NVObjCell::getSheet() {
	return sheet;
}

int NVObjCell::getRow() {
	return row;
}

int NVObjCell::getCol() {
	return col;
}

void NVObjCell::setDependencies(shared_ptr<NumFormatMgr> nf, shared_ptr<Book> b, Sheet* s, int r, int c) {
	setNumFormatMgr(nf);
	setBook(b);
	setSheet(s);
	setRow(r);
	setCol(c);
}

void NVObjCell::setNumFormatMgr(shared_ptr<NumFormatMgr> nf) {
	numFormats = nf;
}

void NVObjCell::setBook(shared_ptr<Book> b) {
	book = b;
}

void NVObjCell::setSheet(Sheet* s) {
	sheet = s;
}

void NVObjCell::setRow(int r) {
	row = r;
}

void NVObjCell::setCol(int c) {
	col = c;
}

/**************************************************************************************************
 **                              CONSTANT TRANSALATION                                           **
 **************************************************************************************************/

// Get the LibXL paper constant for the Omnis paper constant
ErrorType NVObjCell::getErrorTypeConstant(int i) {
	buildErrorTypeLookup();
	return errorTypeLookup[i];
}

// Get the Omnis paper constant for the LibXL paper constant
int NVObjCell::getOmnisErrorTypeConstant(ErrorType c) {
	buildErrorTypeLookup();
	return omnisErrorTypeLookup[c];
}

// Build the maps used by getPaperConstant and getOmnisPaperConstant
//    Allocate the static variables
std::map<int, ErrorType> NVObjCell::errorTypeLookup;
std::map<ErrorType, int> NVObjCell::omnisErrorTypeLookup;

//    One time build of lookup
void NVObjCell::buildErrorTypeLookup() {
	if (errorTypeLookup.size() > 0) // Pseudo-singleton for paper lookup static variables.
		return;
	
	// Build main lookup (based on constant values defined in ExcelFormat.rc)
	errorTypeLookup[1]  = ERRORTYPE_NULL;
	errorTypeLookup[2]  = ERRORTYPE_DIV_0;
	errorTypeLookup[3]  = ERRORTYPE_VALUE;
	errorTypeLookup[4]  = ERRORTYPE_REF;
	errorTypeLookup[5]  = ERRORTYPE_NAME;
	errorTypeLookup[6]  = ERRORTYPE_NUM;
	errorTypeLookup[7]  = ERRORTYPE_NA;
	errorTypeLookup[8]  = ERRORTYPE_NOERROR;
	
	// Build reverse lookup
	std::map<int,ErrorType>::iterator it;
	for ( it=errorTypeLookup.begin() ; it != errorTypeLookup.end(); it++ ) {
		omnisErrorTypeLookup[(*it).second] = (*it).first;
	}
	
	return;
}

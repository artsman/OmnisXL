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

#include "FormatObj.he"
#include "FontObj.he"
#include "Constants.he"

using namespace OmnisTools;
using namespace LibXLConstants;
using namespace libxl;
using boost::shared_ptr;

/**************************************************************************************************
 **                       CONSTRUCTORS / DESTRUCTORS                                             **
 **************************************************************************************************/

NVObjFormat::NVObjFormat(qobjinst qo, tThreadData *pThreadData) : NVObjBase(qo)
{ }

NVObjFormat::~NVObjFormat()
{ }

void NVObjFormat::copy( NVObjFormat* pObj ) {
	// Copy in super class (This does *this = *pObj)
	NVObjBase::copy(pObj);
	
	// Copy the contents of the object into the new object
	book = pObj->book;
	fmt = pObj->fmt;
}

/**************************************************************************************************
 **                              PROPERTY DECLERATION                                            **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project it is also used as the Unique ID.
const static qshort cFMTPropertyLocked                 = 5100,
                    cFMTPropertyHidden                 = 5101,
                    cFMTPropertyFont                   = 5102,
                    cFMTPropertyNumFormat              = 5103,
                    cFMTPropertyAlignH                 = 5104,
                    cFMTPropertyAlignV                 = 5105,
                    cFMTPropertyWrap                   = 5106,
                    cFMTPropertyRotation               = 5107,
                    cFMTPropertyIndent                 = 5108,
                    cFMTPropertyShrinkToFit            = 5109,
                    cFMTPropertyBorderLeft             = 5110,
                    cFMTPropertyBorderRight            = 5111,
                    cFMTPropertyBorderTop              = 5112,
                    cFMTPropertyBorderBottom           = 5113,
                    cFMTPropertyBorderLeftColor        = 5114,
                    cFMTPropertyBorderRightColor       = 5115,
                    cFMTPropertyBorderTopColor         = 5116,
                    cFMTPropertyBorderBottomColor      = 5117,
                    cFMTPropertyBorderDiagnol          = 5118,
                    cFMTPropertyBorderDiagnolColor     = 5119,
                    cFMTPropertyFillPattern            = 5120,
                    cFMTPropertyPatternForegroundColor = 5121,
                    cFMTPropertyPatternBackgroundColor = 5122;


/**************************************************************************************************
 **                               METHOD DECLERATION                                             **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project is also used as the Unique ID.
const static qshort cFMTMethodSetBorder      = 5000,
                    cFMTMethodSetBorderColor = 5001,
                    cFMTMethodSelf           = 5002;

/**************************************************************************************************
 **                                 INSTANCE METHODS                                             **
 **************************************************************************************************/

// Call a method
qlong NVObjFormat::methodCall( tThreadData* pThreadData )
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
		case cFMTMethodSetBorder:
			pThreadData->mCurMethodName = "$setBorder";
			result = methodSetBorder(pThreadData, paramCount);
			break;
		case cFMTMethodSetBorderColor:
			pThreadData->mCurMethodName = "$setBorderColor";
			result = methodSetBorderColor(pThreadData, paramCount);
			break;
        case cFMTMethodSelf:
			pThreadData->mCurMethodName = "$self";
			result = methodSelf(pThreadData, paramCount);
			break;
	}
	
	// Check for errors to notify Omnis developer in $error
	callErrorMethod( pThreadData, result );
	
	return 0L;
}

/**************************************************************************************************
 **                                PROPERTIES                                                    **
 **************************************************************************************************/

// Assignability of properties
qlong NVObjFormat::canAssignProperty( tThreadData* pThreadData, qlong propID ) {
	switch (propID) {
		case cFMTPropertyLocked:
		case cFMTPropertyHidden:
		case cFMTPropertyFont:
		case cFMTPropertyNumFormat:
		case cFMTPropertyAlignH:
		case cFMTPropertyAlignV:
		case cFMTPropertyWrap:
		case cFMTPropertyRotation:
		case cFMTPropertyIndent:
		case cFMTPropertyShrinkToFit:
		case cFMTPropertyBorderLeft:
		case cFMTPropertyBorderRight:
		case cFMTPropertyBorderTop:
		case cFMTPropertyBorderBottom:
		case cFMTPropertyBorderLeftColor:
		case cFMTPropertyBorderRightColor:
		case cFMTPropertyBorderTopColor:
		case cFMTPropertyBorderBottomColor:
		case cFMTPropertyBorderDiagnol:
		case cFMTPropertyBorderDiagnolColor:
		case cFMTPropertyFillPattern:
		case cFMTPropertyPatternForegroundColor:
		case cFMTPropertyPatternBackgroundColor:
			return qtrue;
		default:
			return qfalse;
	}
}

// Method to retrieve a property of the object
qlong NVObjFormat::getProperty( tThreadData* pThreadData ) 
{
	EXTfldval fValReturn;
	if (!checkDependencies()) {
		ECOaddParam(pThreadData->mEci, &fValReturn);
		return qtrue;
	}
	
	int intAssign;
	qlong constAssign;
	Font* fontAssign;
	NVObjFont* omFont;

	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cFMTPropertyLocked:
			getEXTFldValFromBool(fValReturn, fmt->locked() );
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyHidden:
			getEXTFldValFromBool(fValReturn, fmt->hidden() );
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyFont:
			fontAssign = fmt->font();
			if (fontAssign) {
				omFont = createNVObj<NVObjFont>(pThreadData);
				if (omFont) {
					omFont->setDependencies(book, fontAssign);
					omFont->setErrorInstance(mErrorInst);
					getEXTFldValForObj<NVObjFont>(fValReturn, omFont);
					ECOaddParam(pThreadData->mEci, &fValReturn);
				}
			}
			break;
		case cFMTPropertyNumFormat:
			intAssign = getOmnisNumFormatConstant(fmt->numFormat());
			if (intAssign > 0) {
				getEXTFldValFromConstant(fValReturn, NUMFORMAT_CONST_START+(intAssign-1), kConstResourcePrefix);
			} else {
				intAssign = fmt->numFormat();
				getEXTFldValFromInt(fValReturn, intAssign);
			}
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyAlignH:
			constAssign = getOmnisAlignHConstant(fmt->alignH());
			getEXTFldValFromConstant(fValReturn, HALIGN_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyAlignV:
			constAssign = getOmnisAlignVConstant(fmt->alignV());
			getEXTFldValFromConstant(fValReturn, VALIGN_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyWrap:
			getEXTFldValFromBool(fValReturn, fmt->wrap() );
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyRotation:
			getEXTFldValFromInt(fValReturn, fmt->rotation() );
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyIndent:
			getEXTFldValFromInt(fValReturn, fmt->indent() );
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyShrinkToFit:
			getEXTFldValFromBool(fValReturn, fmt->shrinkToFit() );
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderLeft:
			constAssign = getOmnisBorderStyleConstant( fmt->borderLeft() );
			getEXTFldValFromConstant(fValReturn, BORDER_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderRight:
			constAssign = getOmnisBorderStyleConstant( fmt->borderRight() );
			getEXTFldValFromConstant(fValReturn, BORDER_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderTop:
			constAssign = getOmnisBorderStyleConstant( fmt->borderTop() );
			getEXTFldValFromConstant(fValReturn, BORDER_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderBottom:
			constAssign = getOmnisBorderStyleConstant( fmt->borderBottom() );
			getEXTFldValFromConstant(fValReturn, BORDER_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderLeftColor:
			constAssign = getOmnisColorConstant( fmt->borderLeftColor() );
			getEXTFldValFromConstant(fValReturn, COLOR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderRightColor:
			constAssign = getOmnisColorConstant( fmt->borderRightColor() );
			getEXTFldValFromConstant(fValReturn, COLOR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderTopColor:
			constAssign = getOmnisColorConstant( fmt->borderTopColor() );
			getEXTFldValFromConstant(fValReturn, COLOR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderBottomColor:
			constAssign = getOmnisColorConstant( fmt->borderBottomColor() );
			getEXTFldValFromConstant(fValReturn, COLOR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderDiagnol:
			constAssign = getOmnisBorderDiagonalConstant( fmt->borderDiagonal() );
			getEXTFldValFromConstant(fValReturn, DIAGONAL_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyBorderDiagnolColor:
			constAssign = getOmnisColorConstant( fmt->borderDiagonalColor() );
			getEXTFldValFromConstant(fValReturn, COLOR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyFillPattern:
			constAssign = getOmnisFillPatternConstant( fmt->fillPattern() );
			getEXTFldValFromConstant(fValReturn, FILLPATTERN_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyPatternForegroundColor:
			constAssign = getOmnisColorConstant( fmt->patternForegroundColor() );
			getEXTFldValFromConstant(fValReturn, COLOR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFMTPropertyPatternBackgroundColor:
			constAssign = getOmnisColorConstant( fmt->patternBackgroundColor() );
			getEXTFldValFromConstant(fValReturn, COLOR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;	       
	}

	return 1L;
}

// Helper method for getProperty to determine if a number format is a constant or a custom number format
bool isNumFormatConst(int numFormat) {
	switch (numFormat) {
		case NUMFORMAT_GENERAL:
		case NUMFORMAT_NUMBER:
		case NUMFORMAT_NUMBER_D2:
		case NUMFORMAT_NUMBER_SEP:
		case NUMFORMAT_NUMBER_SEP_D2:
		case NUMFORMAT_CURRENCY_NEGBRA:
		case NUMFORMAT_CURRENCY_NEGBRARED:
		case NUMFORMAT_CURRENCY_D2_NEGBRA:
		case NUMFORMAT_CURRENCY_D2_NEGBRARED:
		case NUMFORMAT_PERCENT:
		case NUMFORMAT_PERCENT_D2:
		case NUMFORMAT_SCIENTIFIC_D2:
		case NUMFORMAT_FRACTION_ONEDIG:
		case NUMFORMAT_FRACTION_TWODIG:
		case NUMFORMAT_DATE:
		case NUMFORMAT_CUSTOM_D_MON_YY:
		case NUMFORMAT_CUSTOM_D_MON:
		case NUMFORMAT_CUSTOM_MON_YY:
		case NUMFORMAT_CUSTOM_HMM_AM:
		case NUMFORMAT_CUSTOM_HMMSS_AM:
		case NUMFORMAT_CUSTOM_HMM:
		case NUMFORMAT_CUSTOM_HMMSS:
		case NUMFORMAT_CUSTOM_MDYYYY_HMM:
		case NUMFORMAT_NUMBER_SEP_NEGBRA:
		case NUMFORMAT_NUMBER_SEP_NEGBRARED:
		case NUMFORMAT_NUMBER_D2_SEP_NEGBRA:
		case NUMFORMAT_NUMBER_D2_SEP_NEGBRARED:
		case NUMFORMAT_ACCOUNT:
		case NUMFORMAT_ACCOUNTCUR:
		case NUMFORMAT_ACCOUNT_D2:
		case NUMFORMAT_ACCOUNT_D2_CUR:
		case NUMFORMAT_CUSTOM_MMSS:
		case NUMFORMAT_CUSTOM_H0MMSS:
		case NUMFORMAT_CUSTOM_MMSS0:
		case NUMFORMAT_CUSTOM_000P0E_PLUS0:
		case NUMFORMAT_TEXT:
			return true;
		default:
			return false;
	}
}

// Method to set a property of the object
qlong NVObjFormat::setProperty( tThreadData* pThreadData )
{
	if (!checkDependencies()) {
		return qfalse;
	}
	
	// Retrieve value to set for property, always in first parameter
	EXTfldval fVal;
	if( getParamVar( pThreadData->mEci, 1, fVal) == qfalse ) 
		return qfalse;

	NVObjFont* omFont;
	Font* font;
	int intAssign, intTest;
	bool boolAssign, boolTest;
	AlignH alignHAssign, alignHTest;
	AlignV alignVAssign, alignVTest;
	BorderStyle borderAssign, borderTest;
	BorderDiagonal borderDiagAssign, borderDiagTest;
	Color colorAssign, colorTest;
	FillPattern fillAssign, fillTest;

	// Assign to the appropriate property
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cFMTPropertyLocked:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = fmt->locked();
			if (boolTest != boolAssign) {
				fmt->setLocked(boolAssign);
			}
			break;
		case cFMTPropertyHidden:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = fmt->hidden();
			if (boolTest != boolAssign) {
				fmt->setHidden(boolAssign);
			}
			break;
		case cFMTPropertyFont:
			omFont = getObjForEXTfldval<NVObjFont>(pThreadData, fVal);
			if (omFont) {
				font = omFont->getFont();
				if (font) {
					fmt->setFont(font);
				}
			}	
			break;
		case cFMTPropertyNumFormat:
			if (getType(fVal).valType == fftConstant || getType(fVal).valType == fftCharacter) {
				// Look up constant
				intTest = getIntFromEXTFldVal(fVal, NUMFORMAT_CONST_START, NUMFORMAT_CONST_END);
				if (intTest != -1  && isValidNumFormat(intTest)) {
					intAssign = getNumFormatConstant(intTest);
					if (isNumFormatConst(intAssign)) {
						intTest = fmt->numFormat();
						if (intTest != intAssign) {
							fmt->setNumFormat(intAssign);
						}
					}
				}
			} else {
				// Just use integer
				intAssign = getIntFromEXTFldVal(fVal);
				intTest = fmt->numFormat();
				if (intTest != intAssign) {
					fmt->setNumFormat(intAssign);
				}
			}
			break;
		case cFMTPropertyAlignH:
			intTest = getIntFromEXTFldVal(fVal, HALIGN_CONST_START, HALIGN_CONST_END);
			if (intTest != -1  && isValidAlignH(intTest)) {
				alignHAssign = getAlignHConstant(intTest);
				alignHTest = fmt->alignH();
				if (alignHTest != alignHAssign) {
					fmt->setAlignH(alignHAssign);
				}
			}
			break;
		case cFMTPropertyAlignV:
			intTest = getIntFromEXTFldVal(fVal, VALIGN_CONST_START, VALIGN_CONST_END);
			if (intTest != -1  && isValidAlignV(intTest)) {
				alignVAssign = getAlignVConstant(intTest);
				alignVTest = fmt->alignV();
				if (alignVTest != alignVAssign) {
					fmt->setAlignV(alignVAssign);
				}
			}
			break;
		case cFMTPropertyWrap:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = fmt->wrap();
			if (boolTest != boolAssign) {
				fmt->setWrap(boolAssign);
			}
			break;
		case cFMTPropertyRotation:
			intAssign = getIntFromEXTFldVal(fVal);
			intTest = fmt->rotation();
			if (intTest != intAssign && ((intAssign > 0 && intAssign < 180) || intAssign == 255)) {
				fmt->setRotation(intAssign);
			}
			break;
		case cFMTPropertyIndent:
			intAssign = getIntFromEXTFldVal(fVal);
			intTest = fmt->indent();
			if (intTest != intAssign) {
				fmt->setIndent(intAssign);
			}
			break;
			break;
		case cFMTPropertyShrinkToFit:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = fmt->shrinkToFit();
			if (boolTest != boolAssign) {
				fmt->setShrinkToFit(boolAssign);
			}
			break;
		case cFMTPropertyBorderLeft:
			intTest = getIntFromEXTFldVal(fVal, BORDER_CONST_START, BORDER_CONST_END);
			if (intTest != -1 && isValidBorder(intTest)) {
				borderAssign = getBorderStyleConstant(intTest);
				borderTest = fmt->borderLeft();
				if (borderTest != borderAssign) {
					fmt->setBorderLeft(borderAssign);
				}
			}
			break;
		case cFMTPropertyBorderRight:
			intTest = getIntFromEXTFldVal(fVal, BORDER_CONST_START, BORDER_CONST_END);
			if (intTest != -1 && isValidBorder(intTest)) {
				borderAssign = getBorderStyleConstant(intTest);
				borderTest = fmt->borderRight();
				if (borderTest != borderAssign) {
					fmt->setBorderRight(borderAssign);
				}
			}
			break;
		case cFMTPropertyBorderTop:
			intTest = getIntFromEXTFldVal(fVal, BORDER_CONST_START, BORDER_CONST_END);
			if (intTest != -1 && isValidBorder(intTest)) {
				borderAssign = getBorderStyleConstant(intTest);
				borderTest = fmt->borderTop();
				if (borderTest != borderAssign) {
					fmt->setBorderTop(borderAssign);
				}
			}
			break;
		case cFMTPropertyBorderBottom:
			intTest = getIntFromEXTFldVal(fVal, BORDER_CONST_START, BORDER_CONST_END);
			if (intTest != -1 && isValidBorder(intTest)) {
				borderAssign = getBorderStyleConstant(intTest);
				borderTest = fmt->borderBottom();
				if (borderTest != borderAssign) {
					fmt->setBorderBottom(borderAssign);
				}
			}
			break;
		case cFMTPropertyBorderLeftColor:
			intTest = getIntFromEXTFldVal(fVal, COLOR_CONST_START, COLOR_CONST_END);
			if (intTest != -1 && isValidColor(intTest)) {
				colorAssign = getColorConstant(intTest);
				colorTest = fmt->borderLeftColor();
				if (colorTest != colorAssign) {
					fmt->setBorderLeftColor(colorAssign);
				}
			}
			break;
		case cFMTPropertyBorderRightColor:
			intTest = getIntFromEXTFldVal(fVal, COLOR_CONST_START, COLOR_CONST_END);
			if (intTest != -1 && isValidColor(intTest)) {
				colorAssign = getColorConstant(intTest);
				colorTest = fmt->borderRightColor();
				if (colorTest != colorAssign) {
					fmt->setBorderRightColor(colorAssign);
				}
			}
			break;
		case cFMTPropertyBorderTopColor:
			intTest = getIntFromEXTFldVal(fVal, COLOR_CONST_START, COLOR_CONST_END);
			if (intTest != -1 && isValidColor(intTest)) {
				colorAssign = getColorConstant(intTest);
				colorTest = fmt->borderTopColor();
				if (colorTest != colorAssign) {
					fmt->setBorderTopColor(colorAssign);
				}
			}
			break;
		case cFMTPropertyBorderBottomColor:
			intTest = getIntFromEXTFldVal(fVal, COLOR_CONST_START, COLOR_CONST_END);
			if (intTest != -1 && isValidColor(intTest)) {
				colorAssign = getColorConstant(intTest);
				colorTest = fmt->borderBottomColor();
				if (colorTest != colorAssign) {
					fmt->setBorderBottomColor(colorAssign);
				}
			}
			break;
		case cFMTPropertyBorderDiagnol:
			intTest = getIntFromEXTFldVal(fVal, DIAGONAL_CONST_START, DIAGONAL_CONST_END);
			if (intTest != -1 && isValidBorderDiagonal(intTest)) {
				borderDiagAssign = getBorderDiagonalConstant(intTest);
				borderDiagTest = fmt->borderDiagonal();
				if (borderDiagTest != borderDiagAssign) {
					fmt->setBorderDiagonal(borderDiagAssign);
				}
			}
			break;
		case cFMTPropertyBorderDiagnolColor:
			intTest = getIntFromEXTFldVal(fVal, COLOR_CONST_START, COLOR_CONST_END);
			if (intTest != -1 && isValidColor(intTest)) {
				colorAssign = getColorConstant(intTest);
				colorTest = fmt->borderDiagonalColor();
				if (colorTest != colorAssign) {
					fmt->setBorderDiagonalColor(colorAssign);
				}
			}
			break;
		case cFMTPropertyFillPattern:
			intTest = getIntFromEXTFldVal(fVal, FILLPATTERN_CONST_START, FILLPATTERN_CONST_END);
			if (intTest != -1 && isValidFillPattern(intTest)) {
				fillAssign = getFillPatternConstant(intTest);
				fillTest = fmt->fillPattern();
				if (fillTest != fillAssign) {
					fmt->setFillPattern(fillAssign);
				}
			}
			break;
		case cFMTPropertyPatternForegroundColor:
			intTest = getIntFromEXTFldVal(fVal, COLOR_CONST_START, COLOR_CONST_END);
			if (intTest != -1 && isValidColor(intTest)) {
				colorAssign = getColorConstant(intTest);
				colorTest = fmt->patternForegroundColor();
				if (colorTest != colorAssign) {
					fmt->setPatternForegroundColor(colorAssign);
				}
			}
			break;
		case cFMTPropertyPatternBackgroundColor:
			intTest = getIntFromEXTFldVal(fVal, COLOR_CONST_START, COLOR_CONST_END);
			if (intTest != -1 && isValidColor(intTest)) {
				colorAssign = getColorConstant(intTest);
				colorTest = fmt->patternBackgroundColor();
				if (colorTest != colorAssign) {
					fmt->setPatternBackgroundColor(colorAssign);
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
ECOparam cFormatMethodsParamsTable[] = 
{
	// $setBorder
	5900, fftInteger, 0, 0,
	// $setBorderColor
	5901, fftInteger, 0, 0,
    // $self
	5902, fftBoolean, EXTD_FLAG_PARAMOPT, 0
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
ECOmethodEvent cFormatMethodsTable[] = 
{
	cFMTMethodSetBorder,        cFMTMethodSetBorder,        fftNone,   1, &cFormatMethodsParamsTable[0], 0, 0,
	cFMTMethodSetBorderColor,   cFMTMethodSetBorderColor,   fftNone,   1, &cFormatMethodsParamsTable[1], 0, 0,
    cFMTMethodSelf,             cFMTMethodSelf,             fftObject, 1, &cFormatMethodsParamsTable[2], 0, 0
};

// List of methods in Simple
qlong NVObjFormat::returnMethods(tThreadData* pThreadData)
{
	const qshort cCMethodCount = sizeof(cFormatMethodsTable) / sizeof(ECOmethodEvent);
	
	return ECOreturnMethods( gInstLib, pThreadData->mEci, &cFormatMethodsTable[0], cCMethodCount );
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
ECOproperty cFormatPropertyTable[] = 
{
	cFMTPropertyLocked,                 cFMTPropertyLocked,                 fftBoolean,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFMTPropertyHidden,                 cFMTPropertyHidden,                 fftBoolean,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFMTPropertyFont,                   cFMTPropertyFont,                   fftObject,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFMTPropertyNumFormat,              cFMTPropertyNumFormat,              fftConstant, EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFMTPropertyAlignH,                 cFMTPropertyAlignH,                 fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, HALIGN_CONST_START, HALIGN_CONST_END,
	cFMTPropertyAlignV,                 cFMTPropertyAlignV,                 fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, VALIGN_CONST_START, VALIGN_CONST_END,
	cFMTPropertyWrap,                   cFMTPropertyWrap,                   fftBoolean,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFMTPropertyRotation,               cFMTPropertyRotation,               fftInteger,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFMTPropertyIndent,                 cFMTPropertyIndent,                 fftInteger,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFMTPropertyShrinkToFit,            cFMTPropertyShrinkToFit,            fftBoolean,  EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFMTPropertyBorderLeft,             cFMTPropertyBorderLeft,             fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, BORDER_CONST_START, BORDER_CONST_END,
	cFMTPropertyBorderRight,            cFMTPropertyBorderRight,            fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, BORDER_CONST_START, BORDER_CONST_END,
	cFMTPropertyBorderTop,              cFMTPropertyBorderTop,              fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, BORDER_CONST_START, BORDER_CONST_END,
	cFMTPropertyBorderBottom,           cFMTPropertyBorderBottom,           fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, BORDER_CONST_START, BORDER_CONST_END,
	cFMTPropertyBorderLeftColor,        cFMTPropertyBorderLeftColor,        fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, COLOR_CONST_START, COLOR_CONST_END,
	cFMTPropertyBorderRightColor,       cFMTPropertyBorderRightColor,       fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, COLOR_CONST_START, COLOR_CONST_END,
	cFMTPropertyBorderTopColor,         cFMTPropertyBorderTopColor,         fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, COLOR_CONST_START, COLOR_CONST_END,
	cFMTPropertyBorderBottomColor,      cFMTPropertyBorderBottomColor,      fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, COLOR_CONST_START, COLOR_CONST_END,
	cFMTPropertyBorderDiagnol,          cFMTPropertyBorderDiagnol,          fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, DIAGONAL_CONST_START, DIAGONAL_CONST_END,
	cFMTPropertyBorderDiagnolColor,     cFMTPropertyBorderDiagnolColor,     fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, COLOR_CONST_START, COLOR_CONST_END,
	cFMTPropertyFillPattern,            cFMTPropertyFillPattern,            fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, FILLPATTERN_CONST_START, FILLPATTERN_CONST_END,
	cFMTPropertyPatternForegroundColor, cFMTPropertyPatternForegroundColor, fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, COLOR_CONST_START, COLOR_CONST_END,
	cFMTPropertyPatternBackgroundColor, cFMTPropertyPatternBackgroundColor, fftConstant, EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, COLOR_CONST_START, COLOR_CONST_END,
};

// List of properties in Simple
qlong NVObjFormat::returnProperties( tThreadData* pThreadData )
{
	const qshort propertyCount = sizeof(cFormatPropertyTable) / sizeof(ECOproperty);

	return ECOreturnProperties( gInstLib, pThreadData->mEci, &cFormatPropertyTable[0], propertyCount );
}

/**************************************************************************************************
 **                              OMNIS VISIBLE METHODS                                           **
 **************************************************************************************************/

tResult NVObjFormat::methodSetBorder( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	qlong omBorderConst;
	if ( getParamLong(pThreadData, 1, omBorderConst) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, borderStyleConstant, is required.";
		return ERR_BAD_PARAMS;
	}
	int borderConst = static_cast<int>(omBorderConst);
	
	fmt->setBorder(getBorderStyleConstant(borderConst));
	
	return METHOD_DONE_RETURN;
}

tResult NVObjFormat::methodSetBorderColor( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
	qlong omBorderColor;
	if ( getParamLong(pThreadData, 1, omBorderColor) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, borderColor, is required.";
		return ERR_BAD_PARAMS;
	}
	int borderColor = static_cast<int>(omBorderColor);
	
	fmt->setBorderColor(getColorConstant(borderColor));
	
	return METHOD_DONE_RETURN;
}

// WORKAROUND: Return a copy of the current object.  This is required for setting some properties due to an Omnis bug.
tResult NVObjFormat::methodSelf( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
    
    qbool asObjRef = qtrue;  // Default is to return an object reference
    if ( pParamCount >=1 && getParamBool(pThreadData, 1, asObjRef) != qtrue ) {
            pThreadData->mExtraErrorText = "First parameter, asObjRef, is unrecognized.  Expected boolean.";
            return ERR_BAD_PARAMS;
	}
    
    // Bulid format object
    NVObjFormat* retOmnisFormat = createNVObj<NVObjFormat>(pThreadData);
	retOmnisFormat->setDependencies(book, fmt);
    retOmnisFormat->setErrorInstance(mErrorInst);
	
	// Return format object
	EXTfldval retVal;
	getEXTFldValForObj<NVObjFormat>(retVal, retOmnisFormat, asObjRef);
	ECOaddParam(pThreadData->mEci, &retVal);
    
    return METHOD_DONE_RETURN;
}

/**************************************************************************************************
 **                              INTERNAL      METHODS                                           **
 **************************************************************************************************/

bool NVObjFormat::checkDependencies() {
	if (book && fmt)
		return true;
	else
		return false;
}

shared_ptr<Book> NVObjFormat::getBook() {
	return book;
}

Format* NVObjFormat::getFormat() {
	return fmt;
}

void NVObjFormat::setDependencies(shared_ptr<Book> b, Format* f) {
	setBook(b);
	setFormat(f);
}

void NVObjFormat::setBook(shared_ptr<Book> b) {
	book = b;
}

void NVObjFormat::setFormat(Format* f) {
	fmt = f;
}

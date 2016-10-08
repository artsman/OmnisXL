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

#include "FontObj.he"
#include "Constants.he"

using namespace OmnisTools;
using namespace libxl;
using namespace LibXLConstants;
using boost::shared_ptr;

/**************************************************************************************************
 **                       CONSTRUCTORS / DESTRUCTORS                                             **
 **************************************************************************************************/

NVObjFont::NVObjFont(qobjinst qo, tThreadData *pThreadData) : NVObjBase(qo)
{ }

NVObjFont::~NVObjFont()
{ }

void NVObjFont::copy( NVObjFont* pObj ) {
	// Copy in super class (This does *this = *pObj)
	NVObjBase::copy(pObj);
	
	// Copy the contents of the object into the new object
	book = pObj->book;
	font = pObj->font;
}

/**************************************************************************************************
 **                              PROPERTY DECLERATION                                            **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project it is also used as the Unique ID.
const static qshort cFTPropertySize      = 6100,
                    cFTPropertyItalic    = 6101,
                    cFTPropertyStrikeOut = 6102,
                    cFTPropertyColor     = 6103,
                    cFTPropertyBold      = 6104,
                    cFTPropertyScript    = 6105,
                    cFTPropertyUnderline = 6106,
                    cFTPropertyName      = 6107;

/**************************************************************************************************
 **                               METHOD DECLERATION                                             **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project is also used as the Unique ID.
const static qshort cFTMethodSelf = 6000;

/**************************************************************************************************
 **                                 INSTANCE METHODS                                             **
 **************************************************************************************************/

// Call a method
qlong NVObjFont::methodCall( tThreadData* pThreadData )
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
        case cFTMethodSelf:
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
qlong NVObjFont::canAssignProperty( tThreadData* pThreadData, qlong propID ) {
	switch (propID) {
		case cFTPropertySize:
		case cFTPropertyItalic:
		case cFTPropertyStrikeOut:
		case cFTPropertyColor:
		case cFTPropertyBold:
		case cFTPropertyScript:
		case cFTPropertyUnderline:
		case cFTPropertyName:
			return qtrue;
		default:
			return qfalse;
	}
}

// Method to retrieve a property of the object
qlong NVObjFont::getProperty( tThreadData* pThreadData ) 
{
	EXTfldval fValReturn;
	if (!checkDependencies()) {
		ECOaddParam(pThreadData->mEci, &fValReturn);
		return qtrue;
	}
	qlong constAssign;
	
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cFTPropertySize:
			getEXTFldValFromInt(fValReturn, font->size());
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFTPropertyItalic:
			getEXTFldValFromBool(fValReturn, font->italic());
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFTPropertyStrikeOut:
			getEXTFldValFromBool(fValReturn, font->strikeOut());
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFTPropertyColor:
			constAssign = getOmnisColorConstant( font->color() );
			getEXTFldValFromConstant(fValReturn, COLOR_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFTPropertyBold:
			getEXTFldValFromBool(fValReturn, font->bold());
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFTPropertyScript:
			constAssign = getOmnisScriptConstant( font->script() );
			getEXTFldValFromConstant(fValReturn, SCRIPT_CONST_START+(constAssign-1), kConstResourcePrefix);
			break;
		case cFTPropertyUnderline:
			constAssign = getOmnisUnderlineConstant( font->underline() );
			getEXTFldValFromConstant(fValReturn, UNDERLINE_CONST_START+(constAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;
		case cFTPropertyName:
			getEXTFldValFromWChar(fValReturn, font->name());
			ECOaddParam(pThreadData->mEci, &fValReturn);
			break;	       
	}

	return 1L;
}

// Method to set a property of the object
qlong NVObjFont::setProperty( tThreadData* pThreadData )
{
	if (!checkDependencies()) {
		return qfalse;
	}
	
	// Retrieve value to set for property, always in first parameter
	EXTfldval fVal;
	if( getParamVar( pThreadData->mEci, 1, fVal) == qfalse ) 
		return qfalse;

	int intAssign, intTest;
	bool boolAssign, boolTest;
	std::wstring stringAssign, stringTest;
    const wchar_t* rawStringTest;
	Color colorAssign, colorTest;
	Underline underAssign, underTest;
	Script scriptAssign, scriptTest;
	
	// Assign to the appropriate property
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cFTPropertySize:
			intAssign = getIntFromEXTFldVal(fVal);
			intTest = font->size();
			if (intAssign != intTest) {
				font->setSize( intAssign );
			}
			break;
		case cFTPropertyItalic:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = font->italic();
			if (boolAssign != boolTest) {
				font->setItalic( boolAssign );
			}
			break;
		case cFTPropertyStrikeOut:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = font->strikeOut();
			if (boolAssign != boolTest) {
				font->setStrikeOut( boolAssign );
			}
			break;
		case cFTPropertyColor:
			intTest = getIntFromEXTFldVal(fVal, COLOR_CONST_START, COLOR_CONST_END);
			if (intTest != -1  && isValidColor(intTest)) {
				colorAssign = getColorConstant(intTest);
				colorTest = font->color();
				if (colorAssign != colorTest) {
					font->setColor( colorAssign );
				}
			}
			break;
		case cFTPropertyBold:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = font->bold();
			if (boolAssign != boolTest) {
				font->setBold( boolAssign );
			}
			break;
		case cFTPropertyScript:
			intTest = getIntFromEXTFldVal(fVal, SCRIPT_CONST_START, SCRIPT_CONST_END);
			if (intTest != -1  && isValidScript(intTest)) {
				scriptAssign = getScriptConstant(intTest);
				scriptTest = font->script();
				if (scriptAssign != scriptTest) {
					font->setScript( scriptAssign );
				}
			}
			break;
		case cFTPropertyUnderline:
			intTest = getIntFromEXTFldVal(fVal, UNDERLINE_CONST_START, UNDERLINE_CONST_END);
			if (intTest != -1  && isValidUnderline(intTest)) {
				underAssign = getUnderlineConstant(intTest);
				underTest = font->underline();
				if (underAssign != underTest) {
					font->setUnderline( underAssign );
				}
			}
			break;
		case cFTPropertyName:
			stringAssign = getWStringFromEXTFldVal(fVal);
			rawStringTest = font->name();
            if (rawStringTest) {
                stringTest = rawStringTest;
                if (stringAssign != stringTest) {
                    if( font->setName(stringAssign.c_str() ) == false )
                        return qfalse;
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
ECOparam cFontMethodsParamsTable[] =  {
    // $self
	6900, fftBoolean, EXTD_FLAG_PARAMOPT, 0
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
ECOmethodEvent cFontMethodsTable[] = 
{
    cFTMethodSelf, cFTMethodSelf, fftObject, 1, &cFontMethodsParamsTable[0], 0, 0
};

// List of methods
qlong NVObjFont::returnMethods(tThreadData* pThreadData)
{
	const qshort cCMethodCount = sizeof(cFontMethodsTable) / sizeof(ECOmethodEvent);
	
	return ECOreturnMethods( gInstLib, pThreadData->mEci, &cFontMethodsTable[0], cCMethodCount );
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
ECOproperty cFontPropertyTable[] = 
{
	cFTPropertySize,      cFTPropertySize,      fftInteger,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFTPropertyItalic,    cFTPropertyItalic,    fftBoolean,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFTPropertyStrikeOut, cFTPropertyStrikeOut, fftBoolean,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFTPropertyColor,     cFTPropertyColor,     fftConstant,  EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, COLOR_CONST_START, COLOR_CONST_END,
	cFTPropertyBold,      cFTPropertyBold,      fftBoolean,   EXTD_FLAG_PROPCUSTOM, 0, 0, 0,
	cFTPropertyScript,    cFTPropertyScript,    fftConstant,  EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, SCRIPT_CONST_START, SCRIPT_CONST_END,
	cFTPropertyUnderline, cFTPropertyUnderline, fftConstant,  EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, UNDERLINE_CONST_START, UNDERLINE_CONST_END,
	cFTPropertyName,      cFTPropertyName,      fftCharacter, EXTD_FLAG_PROPCUSTOM, 0, 0, 0
};

// List of properties in Simple
qlong NVObjFont::returnProperties( tThreadData* pThreadData )
{
	const qshort propertyCount = sizeof(cFontPropertyTable) / sizeof(ECOproperty);

	return ECOreturnProperties( gInstLib, pThreadData->mEci, &cFontPropertyTable[0], propertyCount );
}

/**************************************************************************************************
 **                              OMNIS VISIBLE METHODS                                           **
 **************************************************************************************************/

// WORKAROUND: Return a copy of the current object.  This is required for setting some properties due to an Omnis bug.
tResult NVObjFont::methodSelf( OmnisTools::tThreadData* pThreadData, qshort pParamCount ) {
    
    qbool asObjRef = qtrue;  // Default is to return an object reference
    if ( pParamCount >=1 && getParamBool(pThreadData, 1, asObjRef) != qtrue ) {
        pThreadData->mExtraErrorText = "First parameter, asObjRef, is unrecognized.  Expected boolean.";
        return ERR_BAD_PARAMS;
	}
    
    // Bulid font object
    NVObjFont* retOmnisFont = createNVObj<NVObjFont>(pThreadData);
	retOmnisFont->setDependencies(book, font);
    retOmnisFont->setErrorInstance(mErrorInst);
	
	// Return font object
	EXTfldval retVal;
	getEXTFldValForObj<NVObjFont>(retVal, retOmnisFont, asObjRef);
	ECOaddParam(pThreadData->mEci, &retVal);
    
    return METHOD_DONE_RETURN;
}


/**************************************************************************************************
 **                              INTERNAL      METHODS                                           **
 **************************************************************************************************/

bool NVObjFont::checkDependencies() {
	if (book && font)
		return true;
	else
		return false;
}

shared_ptr<Book> NVObjFont::getBook() {
	return book;
}

Font* NVObjFont::getFont() {
	return font;
}

void NVObjFont::setDependencies(shared_ptr<Book> b, Font* f) {
	setBook(b);
	setFont(f);
}

void NVObjFont::setBook(shared_ptr<Book> b) {
	book = b;
}

void NVObjFont::setFont(Font* f) {
	font = f;
}

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
#include "WorksheetObj.he"
#include "FormatObj.he"
#include "CellObj.he"
#include "Constants.he"

using namespace OmnisTools;
using namespace libxl;
using namespace LibXLConstants;
using boost::shared_ptr;

// Constants
static const int MAX_ROWS = 65535, MAX_COLS = 255;

/**************************************************************************************************
 **                       CONSTRUCTORS / DESTRUCTORS                                             **
 **************************************************************************************************/

NVObjWorksheet::NVObjWorksheet(qobjinst qo, tThreadData *pThreadData) : NVObjBase(qo)
{ }

NVObjWorksheet::~NVObjWorksheet()
{ }

void NVObjWorksheet::copy( NVObjWorksheet* pObj ) {
	// Copy in super class (This does *this = *pObj)
	NVObjBase::copy(pObj);
	
	// Copy the contents of the object into the new object
	book = pObj->book;
	sheet = pObj->sheet;
    numFormats = pObj->numFormats;
}

/**************************************************************************************************
 **                              PROPERTY DECLERATION                                            **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project it is also used as the Unique ID.
const static qshort cWSPropertyFirstRow          = 3500,
                    cWSPropertyLastRow           = 3501,
                    cWSPropertyFirstCol          = 3502,
                    cWSPropertyLastCol           = 3503,
                    cWSPropertyGroupSummaryBelow = 3504,
                    cWSPropertyGroupSummaryRight = 3505,
                    cWSPropertyPaper             = 3506,
                    cWSPropertyDisplayGridlines  = 3507,
                    cWSPropertyPrintGridlines    = 3508,
                    cWSPropertyZoom              = 3509,
                    cWSPropertyPrintZoom         = 3510,
                    cWSPropertyLandscape         = 3511,
                    cWSPropertyHeader            = 3512,
                    cWSPropertyHeaderMargin      = 3513,
                    cWSPropertyFooter            = 3514,
                    cWSPropertyFooterMargin      = 3515,
                    cWSPropertyHCenter           = 3516,
                    cWSPropertyVCenter           = 3517,
                    cWSPropertyMarginLeft        = 3518,
                    cWSPropertyMarginRight       = 3519,
                    cWSPropertyMarginTop         = 3520,
                    cWSPropertyMarginBottom      = 3521,
                    cWSPropertyPrintRowCol       = 3522,
                    cWSPropertyProtected         = 3523;

/**************************************************************************************************
 **                               METHOD DECLERATION                                             **
 **************************************************************************************************/

// This is where the resource # of the methods is defined.  In this project is also used as the Unique ID.
const static qshort cWSMethodCell               = 3000,
                    cWSMethodColWidth           = 3001,
                    cWSMethodRowHeight          = 3002,
                    cWSMethodSetCol             = 3003,
                    cWSMethodSetRow             = 3004,
                    cWSMethodGetMerge           = 3005,
                    cWSMethodSetMerge           = 3006,
                    cWSMethodDelMerge           = 3007,
                    cWSMethodSetHorzPageBreak   = 3008,
                    cWSMethodSetVertPageBreak   = 3009,
                    cWSMethodSplit              = 3010,
                    cWSMethodGroupRows          = 3011,
                    cWSMethodGroupCols          = 3012,
                    cWSMethodClear              = 3013,
                    cWSMethodInsertRow          = 3014,
                    cWSMethodInsertCol          = 3015,
                    cWSMethodRemoveRow          = 3016,
                    cWSMethodRemoveCol          = 3017,
                    cWSMethodSetPrintRepeatRows = 3018,
                    cWSMethodSetPrintRepeatCols = 3019,
                    cWSMethodSetPrintArea       = 3020,
                    cWSMethodClearPrintRepeats  = 3021,
                    cWSMethodClearPrintArea     = 3022,
                    cWSMethodSetNamedRange      = 3023,
                    cWSMethodDelNamedRange      = 3024;

/**************************************************************************************************
 **                                 INSTANCE METHODS                                             **
 **************************************************************************************************/

// Call a method
qlong NVObjWorksheet::methodCall( tThreadData* pThreadData )
{
	tResult result = ERR_OK;
	qshort funcId = (qshort)ECOgetId(pThreadData->mEci);
	qshort paramCount = ECOgetParamCount(pThreadData->mEci);

	if (!checkDependencies()) {
		pThreadData->mExtraErrorText = "Internal Error.  Object dependencies are not setup.";
		callErrorMethod( pThreadData, ERR_METHOD_FAILED );
		return qfalse;
	}
	
	switch( funcId )
	{
		case cWSMethodCell:
			pThreadData->mCurMethodName = "$cell";
			result = methodCell(pThreadData, paramCount);
			break;
		case cWSMethodColWidth:
			pThreadData->mCurMethodName = "$colWidth";
			result = methodColWidth(pThreadData, paramCount);
			break;
		case cWSMethodRowHeight:
			pThreadData->mCurMethodName = "$rowHeight";
			result = methodRowHeight(pThreadData, paramCount);
			break;
		case cWSMethodSetCol:
			pThreadData->mCurMethodName = "$setCol";
			result = methodSetCol(pThreadData, paramCount);
			break;
		case cWSMethodSetRow:
			pThreadData->mCurMethodName = "$setRow";
			result = methodSetRow(pThreadData, paramCount);
			break;
		case cWSMethodGetMerge:
			pThreadData->mCurMethodName = "$getMerge";
			result = methodGetMerge(pThreadData, paramCount);
			break;
		case cWSMethodSetMerge:
			pThreadData->mCurMethodName = "$setMerge";
			result = methodSetMerge(pThreadData, paramCount);
			break;
		case cWSMethodDelMerge:
			pThreadData->mCurMethodName = "$delMerge";
			result = methodDelMerge(pThreadData, paramCount);
			break;
		case cWSMethodSetHorzPageBreak:
			pThreadData->mCurMethodName = "$setHorzPageBreak";
			result = methodSetHorzPageBreak(pThreadData, paramCount);
			break;
		case cWSMethodSetVertPageBreak:
			pThreadData->mCurMethodName = "$setVertPageBreak";
			result = methodSetVertPageBreak(pThreadData, paramCount);
			break;
		case cWSMethodSplit:
			pThreadData->mCurMethodName = "$split";
			result = methodSplit(pThreadData, paramCount);
			break;
		case cWSMethodGroupRows:
			pThreadData->mCurMethodName = "$groupRows";
			result = methodGroupRows(pThreadData, paramCount);
			break;
		case cWSMethodGroupCols:
			pThreadData->mCurMethodName = "$groupCols";
			result = methodGroupCols(pThreadData, paramCount);
			break;
		case cWSMethodClear:
			pThreadData->mCurMethodName = "$clear";
			result = methodClear(pThreadData, paramCount);
			break;
		case cWSMethodInsertRow:
			pThreadData->mCurMethodName = "$insertRow";
			result = methodInsertRow(pThreadData, paramCount);
			break;
		case cWSMethodInsertCol:
			pThreadData->mCurMethodName = "$insertCol";
			result = methodInsertCol(pThreadData, paramCount);
			break;
		case cWSMethodRemoveRow:
			pThreadData->mCurMethodName = "$removeRow";
			result = methodRemoveRow(pThreadData, paramCount);
			break;
		case cWSMethodRemoveCol:
			pThreadData->mCurMethodName = "$removeCol";
			result = methodRemoveCol(pThreadData, paramCount);
			break;
		case cWSMethodSetPrintRepeatRows:
			pThreadData->mCurMethodName = "$setPrintRepeatRows";
			result = methodSetPrintRepeatRows(pThreadData, paramCount);
			break;
		case cWSMethodSetPrintRepeatCols:
			pThreadData->mCurMethodName = "$setPrintRepeatCols";
			result = methodSetPrintRepeatCols(pThreadData, paramCount);
			break;
		case cWSMethodSetPrintArea:
			pThreadData->mCurMethodName = "$setPrintArea";
			result = methodSetPrintArea(pThreadData, paramCount);
			break;
		case cWSMethodClearPrintRepeats:
			pThreadData->mCurMethodName = "$clearPrintRepeats";
			result = methodClearPrintRepeats(pThreadData, paramCount);
			break;
		case cWSMethodClearPrintArea:
			pThreadData->mCurMethodName = "$clearPrintArea";
			result = methodClearPrintArea(pThreadData, paramCount);
			break;
		case cWSMethodSetNamedRange:
			pThreadData->mCurMethodName = "$setNamedRange";
			result = methodSetNamedRange(pThreadData, paramCount);
			break;
		case cWSMethodDelNamedRange:
			pThreadData->mCurMethodName = "$delNamedRange";
			result = methodDelNamedRange(pThreadData, paramCount);
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
qlong NVObjWorksheet::canAssignProperty( tThreadData* pThreadData, qlong propID ) {
	switch (propID) {
		case cWSPropertyGroupSummaryBelow:
		case cWSPropertyGroupSummaryRight:
		case cWSPropertyPaper:
		case cWSPropertyDisplayGridlines:
		case cWSPropertyPrintGridlines:
		case cWSPropertyZoom:
		case cWSPropertyPrintZoom:
		case cWSPropertyLandscape:
		case cWSPropertyHeader:
		case cWSPropertyHeaderMargin:
		case cWSPropertyFooter:
		case cWSPropertyFooterMargin:
		case cWSPropertyHCenter:
		case cWSPropertyVCenter:
		case cWSPropertyMarginLeft:
		case cWSPropertyMarginRight:
		case cWSPropertyMarginTop:
		case cWSPropertyMarginBottom:
		case cWSPropertyPrintRowCol:
		case cWSPropertyProtected:
			return qtrue;
		case cWSPropertyFirstRow:
		case cWSPropertyLastRow:
		case cWSPropertyFirstCol:
		case cWSPropertyLastCol:
			return qfalse;
		default:
			return qfalse;
	}
}

// Method to retrieve a property of the object
qlong NVObjWorksheet::getProperty( tThreadData* pThreadData ) 
{
	EXTfldval fValReturn;
	if (!checkDependencies()) {
		ECOaddParam(pThreadData->mEci, &fValReturn);
		return qtrue;
	}
	
	// Generic variables for returning the properties
	int intAssign;
	bool boolAssign;
	double doubleAssign;
	const wchar_t* rawStringTest;
	
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cWSPropertyFirstRow:
			intAssign = sheet->firstRow() + 1; // Index off for Omnis' sake
			getEXTFldValFromInt(fValReturn, intAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyLastRow:
			intAssign = sheet->lastRow() + 1;  // Index off for Omnis' sake
			getEXTFldValFromInt(fValReturn, intAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyFirstCol:
			intAssign = sheet->firstCol() + 1;  // Index off for Omnis' sake
			getEXTFldValFromInt(fValReturn, intAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyLastCol:
			intAssign = sheet->lastCol() + 1;  // Index off for Omnis' sake
			getEXTFldValFromInt(fValReturn, intAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyGroupSummaryBelow:
			boolAssign = sheet->groupSummaryBelow();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyGroupSummaryRight:
			boolAssign = sheet->groupSummaryRight();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyPaper:
			intAssign = getOmnisPaperConstant(sheet->paper());
			getEXTFldValFromConstant(fValReturn, PAPER_CONST_START+(intAssign-1), kConstResourcePrefix);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyDisplayGridlines:
			boolAssign = sheet->displayGridlines();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyPrintGridlines:
			boolAssign = sheet->printGridlines();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyZoom:
			intAssign = sheet->zoom();
			getEXTFldValFromInt(fValReturn, intAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyPrintZoom:
			intAssign = sheet->printZoom();
			getEXTFldValFromInt(fValReturn, intAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyLandscape:
			boolAssign = sheet->landscape();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyHeader:
			rawStringTest = sheet->header();
			if (rawStringTest) {
				// Valid string
				getEXTFldValFromWString(fValReturn, rawStringTest);
			} else {
				getEXTFldValFromString(fValReturn, "");
			}
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyHeaderMargin:
			doubleAssign = sheet->headerMargin();
			getEXTFldValFromDouble(fValReturn, doubleAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyFooter:
			rawStringTest = sheet->footer();
			if (rawStringTest) {
				// Valid string
				getEXTFldValFromWChar(fValReturn, rawStringTest);
			} else {
				getEXTFldValFromString(fValReturn, "");
			}
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyFooterMargin:
			doubleAssign = sheet->footerMargin();
			getEXTFldValFromDouble(fValReturn, doubleAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyHCenter:
			boolAssign = sheet->hCenter();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyVCenter:
			boolAssign = sheet->vCenter();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyMarginLeft:
			doubleAssign = sheet->marginLeft();
			getEXTFldValFromDouble(fValReturn, doubleAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyMarginRight:
			doubleAssign = sheet->marginRight();
			getEXTFldValFromDouble(fValReturn, doubleAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyMarginTop:
			doubleAssign = sheet->marginTop();
			getEXTFldValFromDouble(fValReturn, doubleAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyMarginBottom:
			doubleAssign = sheet->marginBottom();
			getEXTFldValFromDouble(fValReturn, doubleAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyPrintRowCol:
			boolAssign = sheet->printRowCol();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	
		case cWSPropertyProtected:
			boolAssign = sheet->protect();
			getEXTFldValFromBool(fValReturn, boolAssign);
			ECOaddParam(pThreadData->mEci, &fValReturn); // Return to caller
			break;	       
	}

	return 1L;
}

// Method to set a property of the object
qlong NVObjWorksheet::setProperty( tThreadData* pThreadData )
{
	if (!checkDependencies()) {
		return qfalse;
	}
	
	// Retrieve value to set for property, always in first parameter
	EXTfldval fVal;
	if( getParamVar( pThreadData->mEci, 1, fVal) == qfalse ) 
		return qfalse;
	
	// Generic variables for testing and setting the properties
	int intAssign, intTest;
	bool boolAssign, boolTest;
	double doubleAssign, doubleTest;
    std::wstring stringAssign, stringTest;
	const wchar_t* rawStringTest;
	
	// Assign to the appropriate property
	// All assignments are protected so that libXL is only called when something changes.
	// This is important since Omnis "sets" variables repeatedly when viewed by instance viewer.
	qlong propID = ECOgetId( pThreadData->mEci );
	switch( propID ) {
		case cWSPropertyGroupSummaryBelow:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->groupSummaryBelow();
			if( boolAssign != boolTest)
				sheet->setGroupSummaryBelow(boolAssign);
			break;	
		case cWSPropertyGroupSummaryRight:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->groupSummaryRight();
			if( boolAssign != boolTest)
				sheet->setGroupSummaryRight(boolAssign);
			break;	
		case cWSPropertyPaper:
			intAssign = getIntFromEXTFldVal(fVal, PAPER_CONST_START, PAPER_CONST_END);
			if (intAssign != -1  && isValidPaper(intAssign)) {
				intTest = getOmnisPaperConstant(sheet->paper());
				if (intAssign != intTest)
					sheet->setPaper(getPaperConstant(intAssign));
			}
			break;	
		case cWSPropertyDisplayGridlines:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->displayGridlines();
			if( boolAssign != boolTest)
				sheet->setDisplayGridlines(boolAssign);
			break;	
		case cWSPropertyPrintGridlines:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->printGridlines();
			if( boolAssign != boolTest)
				sheet->setPrintGridlines(boolAssign);
			break;	
		case cWSPropertyZoom:
			intAssign = getIntFromEXTFldVal(fVal);
			intTest = sheet->zoom();
			if (intAssign != intTest)
				sheet->setZoom(intAssign);
			break;	
		case cWSPropertyPrintZoom:
			intAssign = getIntFromEXTFldVal(fVal);
			intTest = sheet->printZoom();
			if (intAssign != intTest)
				sheet->setPrintZoom(intAssign);
			break;	
		case cWSPropertyLandscape:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->landscape();
			if( boolAssign != boolTest)
				sheet->setLandscape(boolAssign);
			break;	
		case cWSPropertyHeader:
			stringAssign = getWStringFromEXTFldVal(fVal);
			rawStringTest = sheet->header();
			if (rawStringTest)
				stringTest = rawStringTest;
			else
				stringTest = L"";

			if (stringAssign != stringTest)
				sheet->setHeader(stringAssign.c_str(), sheet->headerMargin());
			
			break;	
		case cWSPropertyHeaderMargin:
			doubleAssign = getIntFromEXTFldVal(fVal);
			doubleTest = sheet->headerMargin();
			rawStringTest = sheet->header();
			if (rawStringTest)
				stringTest = rawStringTest;
			else
				stringTest = L"";
			
			if (doubleAssign != doubleTest)
				sheet->setHeader(stringAssign.c_str(), doubleAssign);
			break;
		case cWSPropertyFooter:
			stringAssign = getWStringFromEXTFldVal(fVal);
			rawStringTest = sheet->footer();
			if (rawStringTest)
				stringTest = rawStringTest;
			else
				stringTest = L"";
			
			if (stringAssign != stringTest)
				sheet->setFooter(stringAssign.c_str(), sheet->footerMargin());
			
			break;	
		case cWSPropertyFooterMargin:
			doubleAssign = getIntFromEXTFldVal(fVal);
			doubleTest = sheet->footerMargin();
			rawStringTest = sheet->footer();
			if (rawStringTest)
				stringTest = rawStringTest;
			else
				stringTest = L"";
			
			if (doubleAssign != doubleTest)
				sheet->setFooter(stringAssign.c_str(), doubleAssign);
			
			break;	
		case cWSPropertyHCenter:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->hCenter();
			if( boolAssign != boolTest)
				sheet->setHCenter(boolAssign);
			break;	
		case cWSPropertyVCenter:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->vCenter();
			if( boolAssign != boolTest)
				sheet->setVCenter(boolAssign);
			break;	
		case cWSPropertyMarginLeft:
			doubleAssign = getDoubleFromEXTFldVal(fVal);
			doubleTest = sheet->marginLeft();
			if (doubleAssign != doubleTest)
				sheet->setMarginLeft(doubleAssign);
			break;	
		case cWSPropertyMarginRight:
			doubleAssign = getDoubleFromEXTFldVal(fVal);
			doubleTest = sheet->marginRight();
			if (doubleAssign != doubleTest)
				sheet->setMarginRight(doubleAssign);
			break;	
		case cWSPropertyMarginTop:
			doubleAssign = getDoubleFromEXTFldVal(fVal);
			doubleTest = sheet->marginTop();
			if (doubleAssign != doubleTest)
				sheet->setMarginTop(doubleAssign);
			break;	
		case cWSPropertyMarginBottom:
			doubleAssign = getDoubleFromEXTFldVal(fVal);
			doubleTest = sheet->marginBottom();
			if (doubleAssign != doubleTest)
				sheet->setMarginBottom(doubleAssign);
			break;	
		case cWSPropertyPrintRowCol:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->printRowCol();
			if( boolAssign != boolTest)
				sheet->setPrintRowCol(boolAssign);
			break;	
		case cWSPropertyProtected:
			boolAssign = getBoolFromEXTFldVal(fVal);
			boolTest = sheet->protect();
			if( boolAssign != boolTest)
				sheet->setProtect(boolAssign);
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
ECOparam cWorksheetMethodsParamsTable[] = 
{
	// $cell
	3800, fftInteger,   0, 0,
	3801, fftInteger,   0, 0,
	// $colWidth
	3802, fftInteger,   0, 0,
	// $rowHeight
	3803, fftInteger,   0, 0,
	// $setCol
	3804, fftInteger,   0, 0,
	3805, fftInteger,   0, 0,
	3806, fftNumber,    0, 0,
	3807, fftObject,    EXTD_FLAG_PARAMOPT, 0,
	3808, fftBoolean,   EXTD_FLAG_PARAMOPT, 0,
	// $setRow
	3809, fftInteger,   0, 0,
	3810, fftNumber,    0, 0,
	3811, fftObject,    EXTD_FLAG_PARAMOPT, 0,
	3812, fftBoolean,   EXTD_FLAG_PARAMOPT, 0,
	// $getMerge
	3813, fftInteger,   0, 0,
	3814, fftInteger,   0, 0,
	3815, fftInteger,   EXTD_FLAG_PARAMALTER, 0,
	3816, fftInteger,   EXTD_FLAG_PARAMALTER, 0,
	3817, fftInteger,   EXTD_FLAG_PARAMALTER, 0,
	3818, fftInteger,   EXTD_FLAG_PARAMALTER, 0,
	// $setMerge
	3819, fftInteger,   0, 0,
	3820, fftInteger,   0, 0,
	3821, fftInteger,   0, 0,
	3822, fftInteger,   0, 0,
	// $delMerge
	3823, fftInteger,   0, 0,
	3824, fftInteger,   0, 0,
	// $setHorzPageBreak
	3825, fftInteger,   0, 0,
	3826, fftBoolean,   EXTD_FLAG_PARAMOPT, 0,
	// $setVertPageBreak
	3827, fftInteger,   0, 0,
	3828, fftBoolean,   EXTD_FLAG_PARAMOPT, 0,
	// $split
	3829, fftInteger,   0, 0,
	3830, fftInteger,   0, 0,
	// $groupRows
	3831, fftInteger,   0, 0,
	3832, fftInteger,   0, 0,
	3833, fftBoolean,   EXTD_FLAG_PARAMOPT, 0,
	// $groupCols
	3834, fftInteger,   0, 0,
	3835, fftInteger,   0, 0,
	3836, fftBoolean,   EXTD_FLAG_PARAMOPT, 0,
	// $clear
	3837, fftInteger,   EXTD_FLAG_PARAMOPT, 0,
	3838, fftInteger,   EXTD_FLAG_PARAMOPT, 0,
	3839, fftInteger,   EXTD_FLAG_PARAMOPT, 0,
	3840, fftInteger,   EXTD_FLAG_PARAMOPT, 0,
	// $insertRow
	3841, fftInteger,   0, 0,
	3842, fftInteger,   0, 0,
	// $insertCol
	3843, fftInteger,   0, 0,
	3844, fftInteger,   0, 0,
	// $removeRow
	3845, fftInteger,   0, 0,
	3846, fftInteger,   0, 0,
	// $removeCol
	3847, fftInteger,   0, 0,
	3848, fftInteger,   0, 0,
	// $setPrintRepeatRows
	3849, fftInteger,   0, 0,
	3850, fftInteger,   0, 0,
	// $setPrintRepeatCols
	3851, fftInteger,   0, 0,
	3852, fftInteger,   0, 0,
	// $setPrintArea
	3853, fftInteger,   0, 0,
	3854, fftInteger,   0, 0,
	3855, fftInteger,   0, 0,
	3856, fftInteger,   0, 0,
	// $setNamedRange
	3857, fftCharacter, 0, 0,
	3858, fftInteger,   0, 0,
	3859, fftInteger,   0, 0,
	3860, fftInteger,   0, 0,
	3861, fftInteger,   0, 0,
	// $delNamedRange
    3862, fftCharacter, 0, 0,
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
ECOmethodEvent cWorksheetMethodsTable[] = 
{
	cWSMethodCell,               cWSMethodCell,               fftObject,  2, &cWorksheetMethodsParamsTable[0],  0, 0,
	cWSMethodColWidth,           cWSMethodColWidth,           fftNumber,  1, &cWorksheetMethodsParamsTable[2],  0, 0,
	cWSMethodRowHeight,          cWSMethodRowHeight,          fftNumber,  1, &cWorksheetMethodsParamsTable[3],  0, 0,
	cWSMethodSetCol,             cWSMethodSetCol,             fftBoolean, 5, &cWorksheetMethodsParamsTable[4],  0, 0,
	cWSMethodSetRow,             cWSMethodSetRow,             fftBoolean, 4, &cWorksheetMethodsParamsTable[9],  0, 0,
	cWSMethodGetMerge,           cWSMethodGetMerge,           fftBoolean, 6, &cWorksheetMethodsParamsTable[13], 0, 0,
	cWSMethodSetMerge,           cWSMethodSetMerge,           fftBoolean, 4, &cWorksheetMethodsParamsTable[19], 0, 0,
	cWSMethodDelMerge,           cWSMethodDelMerge,           fftBoolean, 2, &cWorksheetMethodsParamsTable[23], 0, 0,
	cWSMethodSetHorzPageBreak,   cWSMethodSetHorzPageBreak,   fftBoolean, 2, &cWorksheetMethodsParamsTable[25], 0, 0,
	cWSMethodSetVertPageBreak,   cWSMethodSetVertPageBreak,   fftBoolean, 2, &cWorksheetMethodsParamsTable[27], 0, 0,
	cWSMethodSplit,              cWSMethodSplit,              fftNone,    2, &cWorksheetMethodsParamsTable[29], 0, 0,
	cWSMethodGroupRows,          cWSMethodGroupRows,          fftBoolean, 3, &cWorksheetMethodsParamsTable[31], 0, 0,
	cWSMethodGroupCols,          cWSMethodGroupCols,          fftBoolean, 3, &cWorksheetMethodsParamsTable[34], 0, 0,
	cWSMethodClear,              cWSMethodClear,              fftNone,    4, &cWorksheetMethodsParamsTable[37], 0, 0,
	cWSMethodInsertRow,          cWSMethodInsertRow,          fftBoolean, 2, &cWorksheetMethodsParamsTable[41], 0, 0,
	cWSMethodInsertCol,          cWSMethodInsertCol,          fftBoolean, 2, &cWorksheetMethodsParamsTable[43], 0, 0,
	cWSMethodRemoveRow,          cWSMethodRemoveRow,          fftBoolean, 2, &cWorksheetMethodsParamsTable[45], 0, 0,
	cWSMethodRemoveCol,          cWSMethodRemoveCol,          fftBoolean, 2, &cWorksheetMethodsParamsTable[47], 0, 0,
	cWSMethodSetPrintRepeatRows, cWSMethodSetPrintRepeatRows, fftNone,    2, &cWorksheetMethodsParamsTable[49], 0, 0,
	cWSMethodSetPrintRepeatCols, cWSMethodSetPrintRepeatCols, fftNone,    2, &cWorksheetMethodsParamsTable[51], 0, 0,
	cWSMethodSetPrintArea,       cWSMethodSetPrintArea,       fftNone,    4, &cWorksheetMethodsParamsTable[53], 0, 0,
	cWSMethodClearPrintRepeats,  cWSMethodClearPrintRepeats,  fftNone,    0,                                 0, 0, 0,
	cWSMethodClearPrintArea,     cWSMethodClearPrintArea,     fftNone,    0,                                 0, 0, 0,
	cWSMethodSetNamedRange,      cWSMethodSetNamedRange,      fftBoolean, 5, &cWorksheetMethodsParamsTable[57], 0, 0,
	cWSMethodDelNamedRange,      cWSMethodDelNamedRange,      fftBoolean, 1, &cWorksheetMethodsParamsTable[62], 0, 0,
};

// List of methods in Simple
qlong NVObjWorksheet::returnMethods(tThreadData* pThreadData)
{
	const qshort cWSMethodCount = sizeof(cWorksheetMethodsTable) / sizeof(ECOmethodEvent);
	
	return ECOreturnMethods( gInstLib, pThreadData->mEci, &cWorksheetMethodsTable[0], cWSMethodCount );
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
ECOproperty cWorksheetPropertyTable[] = 
{
	cWSPropertyFirstRow,          cWSPropertyFirstRow,          fftInteger,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyLastRow,           cWSPropertyLastRow,           fftInteger,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyFirstCol,          cWSPropertyFirstCol,          fftInteger,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyLastCol,           cWSPropertyLastCol,           fftInteger,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyGroupSummaryBelow, cWSPropertyGroupSummaryBelow, fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyGroupSummaryRight, cWSPropertyGroupSummaryRight, fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyPaper,             cWSPropertyPaper,             fftConstant,  EXTD_FLAG_PROPCUSTOM|EXTD_FLAG_EXTCONSTANT, 0, PAPER_CONST_START, PAPER_CONST_END,
	cWSPropertyDisplayGridlines,  cWSPropertyDisplayGridlines,  fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyPrintGridlines,    cWSPropertyPrintGridlines,    fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyZoom,              cWSPropertyZoom,              fftInteger,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyPrintZoom,         cWSPropertyPrintZoom,         fftInteger,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyLandscape,         cWSPropertyLandscape,         fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyHeader,            cWSPropertyHeader,            fftCharacter, EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyHeaderMargin,      cWSPropertyHeaderMargin,      fftNumber,    EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyFooter,            cWSPropertyFooter,            fftCharacter, EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyFooterMargin,      cWSPropertyFooterMargin,      fftNumber,    EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyHCenter,           cWSPropertyHCenter,           fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyVCenter,           cWSPropertyVCenter,           fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyMarginLeft,        cWSPropertyMarginLeft,        fftNumber,    EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyMarginRight,       cWSPropertyMarginRight,       fftNumber,    EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyMarginTop,         cWSPropertyMarginTop,         fftNumber,    EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyMarginBottom,      cWSPropertyMarginBottom,      fftNumber,    EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyPrintRowCol,       cWSPropertyPrintRowCol,       fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0,
	cWSPropertyProtected,         cWSPropertyProtected,         fftBoolean,   EXTD_FLAG_PROPCUSTOM,  0, 0, 0
};

// List of properties in Simple
qlong NVObjWorksheet::returnProperties( tThreadData* pThreadData )
{
	const qshort propertyCount = sizeof(cWorksheetPropertyTable) / sizeof(ECOproperty);

	return ECOreturnProperties( gInstLib, pThreadData->mEci, &cWorksheetPropertyTable[0], propertyCount );
}

/**************************************************************************************************
 **                              OMNIS VISIBLE METHODS                                           **
 **************************************************************************************************/

tResult NVObjWorksheet::methodCell( tThreadData* pThreadData, qshort pParamCount ) {
	// Since Cell is only an object within Omnis, there is no call to libXL
	qlong omRow;
	int row, col;
	EXTfldval colVal;
	
	// Read parameters.  
	// Param 1: Row
	if ( getParamLong(pThreadData, 1, omRow) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, row, is required.";
		return ERR_BAD_PARAMS;
	} 
	row = static_cast<int>(omRow) - 1;  // Index off 0 for LibXL's sake
	
	// Param 2: Column
	if ( getParamVar(pThreadData, 2, colVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, col, is required.";
		return ERR_BAD_PARAMS;
	} 
	col = NVObjWorkbook::getColNumber(colVal) - 1; // Index off 0 for LibXL's sake
	if(col < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}

	// Create omnis representation of sheet and add neccesary pointers
	NVObjCell* retCell = createNVObj<NVObjCell>(pThreadData);
	retCell->setDependencies(numFormats, book, sheet, row, col);
	retCell->setErrorInstance(mErrorInst);
	
	// Return Cell object to caller
	EXTfldval retVal;
	getEXTFldValForObj<NVObjCell>(retVal, retCell);
	ECOaddParam(pThreadData->mEci, &retVal); 
	
	return METHOD_DONE_RETURN;
}

// Set column width
tResult NVObjWorksheet::methodColWidth( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: Column
	EXTfldval colVal;
	if ( getParamVar(pThreadData, 1, colVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, col, is required.";
		return ERR_BAD_PARAMS;
	} 
	int col = NVObjWorkbook::getColNumber(colVal) - 1; // Index off 0 for LibXL's sake
	if(col < 0) {
		pThreadData->mExtraErrorText = "Unknown first parameter.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL 
	double retWidth = sheet->colWidth(col);
	
	// Return
	EXTfldval retVal;
	getEXTFldValFromDouble(retVal, retWidth);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Set row height
tResult NVObjWorksheet::methodRowHeight( tThreadData* pThreadData, qshort pParamCount ) {
  
	// Param 1: Row
	qlong omRow;
	if ( getParamLong(pThreadData, 1, omRow) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, row, is required.";
		return ERR_BAD_PARAMS;
	} 
	int row = static_cast<int>(omRow) - 1; // Index off 0 for LibXL's sake
	
	// Call LibXL 
	double retHeight = sheet->rowHeight(row);
	
	// Return
	EXTfldval retVal;
	getEXTFldValFromDouble(retVal, retHeight);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Set column width and format for all columns between colFirst and colLast
tResult NVObjWorksheet::methodSetCol( tThreadData* pThreadData, qshort pParamCount ) {
	// Param 1: First Column
	EXTfldval colFirstVal;
	if ( getParamVar(pThreadData, 1, colFirstVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, colFirst, is required.";
		return ERR_BAD_PARAMS;
	} 
	int colFirst = NVObjWorkbook::getColNumber(colFirstVal) - 1; // Index off 0 for LibXL's sake
	if(colFirst < 0) {
		pThreadData->mExtraErrorText = "Unknown first parameter.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 2: Last Column
	EXTfldval colLastVal;
	if ( getParamVar(pThreadData, 2, colLastVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, colLast, is required.";
		return ERR_BAD_PARAMS;
	} 
	int colLast = NVObjWorkbook::getColNumber(colLastVal) - 1; // Index off 0 for LibXL's sake
	if(colLast < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 3: Width
	EXTfldval widthVal;
	if ( getParamVar(pThreadData, 3, widthVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Third parameter, width, is required.";
		return ERR_BAD_PARAMS;
	}
	double width = getDoubleFromEXTFldVal(widthVal);
	
	// Param 4: (Optional) Format
	EXTfldval formatVal;
	NVObjFormat* omFormat;
	Format* fmt = 0;
	if ( getParamVar(pThreadData, 4, formatVal) == qtrue ) {
		omFormat = getObjForEXTfldval<NVObjFormat>(pThreadData, formatVal);
		if (!omFormat) {
			pThreadData->mExtraErrorText = "Fourth parameter, format, is unrecognized.";
			return ERR_BAD_PARAMS;
		}
		fmt = omFormat->getFormat(); // Retrieve
	}
	
	// Param 5: (Optional) Hidden
	qbool omHidden;
	bool hidden;
	bool passedHidden = false;
	if ( getParamBool(pThreadData, 5, omHidden) == qtrue ) {
		hidden = getBoolFromQBool(omHidden);
		passedHidden = true;
	}
	
	// Call LibXL depending on which parameters were passed
	bool retBool;
	if (passedHidden) {
		retBool = sheet->setCol(colFirst, colLast, width, fmt, hidden);
	} else if (fmt) {
		retBool = sheet->setCol(colFirst, colLast, width, fmt);
	} else {
		retBool = sheet->setCol(colFirst, colLast, width);
	}
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Set the height for a given row
tResult NVObjWorksheet::methodSetRow( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: Row
	qlong omRow;
	if ( getParamLong(pThreadData, 1, omRow) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, row, is required.";
		return ERR_BAD_PARAMS;
	} 
	int row = static_cast<int>(omRow) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Height
	EXTfldval heightVal;
	if ( getParamVar(pThreadData, 2, heightVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, height, is required.";
		return ERR_BAD_PARAMS;
	}
	double height = getDoubleFromEXTFldVal(heightVal);
	
	// Param 3: (Optional) Format
	EXTfldval formatVal;
	NVObjFormat* omFormat;
	Format* fmt;
	if ( getParamVar(pThreadData, 3, formatVal) == qtrue ) {
		omFormat = getObjForEXTfldval<NVObjFormat>(pThreadData, formatVal);
		if (!omFormat) {
			pThreadData->mExtraErrorText = "Third parameter, format, is unrecognized.";
			return ERR_BAD_PARAMS;
		}
		fmt = omFormat->getFormat(); // Retrieve
	}
	
	// Param 4: (Optional) Hidden
	qbool omHidden;
	bool hidden;
	bool passedHidden = false;
	if ( getParamBool(pThreadData, 4, omHidden) == qtrue ) {
		hidden = getBoolFromQBool(omHidden);
		passedHidden = true;
	}
	
	// Call LibXL depending on which parameters were passed
	bool retBool;
	if (passedHidden) {
		retBool = sheet->setRow(row, height, fmt, hidden);
	} else if (fmt) {
		retBool = sheet->setRow(row, height, fmt);
	} else {
		retBool = sheet->setRow(row, height);
	}
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Gets merged cells for cell at row, col. Result is written in rowFirst, rowLast, colFirst, colLast. Returns false if error occurs. 
tResult NVObjWorksheet::methodGetMerge( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: Row
	qlong omRow;
	if ( getParamLong(pThreadData, 1, omRow) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, row, is required.";
		return ERR_BAD_PARAMS;
	} 
	int row = static_cast<int> (omRow) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Col
	EXTfldval colVal;
	if ( getParamVar(pThreadData, 2, colVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, col, is required.";
		return ERR_BAD_PARAMS;
	} 
	int col = NVObjWorkbook::getColNumber(colVal) - 1; // Index off 0 for LibXL's sake
	if(col < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 3-6: Return values for first row, last row, first col and last col.
	EXTfldval rowFirstVal, rowLastVal, colFirstVal, colLastVal;
	if( getParamVar(pThreadData, 3, rowFirstVal) != qtrue ||
	    getParamVar(pThreadData, 4, rowLastVal)  != qtrue ||
	    getParamVar(pThreadData, 5, colFirstVal) != qtrue ||
	    getParamVar(pThreadData, 6, colLastVal)  != qtrue) 
	{
		pThreadData->mExtraErrorText = "Insufficient parameters.  rowFirst, rowLast, colFirst and colLast must all be passed.";
		return ERR_BAD_PARAMS;
	}	
	
	// Call LibXL
	int rowFirst, rowLast, colFirst, colLast;
	bool retBool = sheet->getMerge(row, col, &rowFirst, &rowLast, &colFirst, &colLast);
	
	if (retBool) {
		// Return values indexed off of 1, not 0.
		rowFirst += 1; rowLast += 1;
		colFirst += 1; colLast += 1;
		
		// Populate passed in values and notify Omnis that the values have changed
		getEXTFldValFromInt(rowFirstVal, rowFirst);
		ECOsetParameterChanged( pThreadData->mEci, 3 );
		
		getEXTFldValFromInt(rowLastVal, rowLast);
		ECOsetParameterChanged( pThreadData->mEci, 4 );
		
		getEXTFldValFromInt(colFirstVal, colFirst);
		ECOsetParameterChanged( pThreadData->mEci, 5 );
		
		getEXTFldValFromInt(colLastVal, colLast);
		ECOsetParameterChanged( pThreadData->mEci, 6 );
	}
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Sets merged cells for range: rowFirst - rowLast, colFirst - colLast. Returns false if error occurs.
tResult NVObjWorksheet::methodSetMerge( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	qlong omRowFirst, omRowLast;
	EXTfldval colFirstVal, colLastVal; 
	int rowFirst, rowLast, colFirst, colLast;
	
	// Param 1: First row to merge
	if( getParamLong(pThreadData, 1, omRowFirst) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, rowFirst, is required.";
		return ERR_BAD_PARAMS;
	}
	rowFirst = static_cast<int> (omRowFirst) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Last row to merge
	if( getParamLong(pThreadData, 2, omRowLast) != qtrue) {
		pThreadData->mExtraErrorText = "Second parameter, rowLast, is required.";
		return ERR_BAD_PARAMS;
	}
	rowLast = static_cast<int> (omRowLast) - 1; // Index off 0 for LibXL's sake
	
	// Param 3: First col to merge
	if ( getParamVar(pThreadData, 3, colFirstVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Third parameter, colFirst, is required.";
		return ERR_BAD_PARAMS;
	} 
	colFirst = NVObjWorkbook::getColNumber(colFirstVal) - 1; // Index off 0 for LibXL's sake
	if(colFirst < 0) {
		pThreadData->mExtraErrorText = "Unknown third parameter, colFirst.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 4: Last col to merge
	if ( getParamVar(pThreadData, 4, colLastVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Fourth parameter, colLast, is required.";
		return ERR_BAD_PARAMS;
	} 
	colLast = NVObjWorkbook::getColNumber(colLastVal) - 1; // Index off 0 for LibXL's sake
	if(colLast < 0) {
		pThreadData->mExtraErrorText = "Unknown fourth parameter, colLast.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	bool retBool = sheet->setMerge(rowFirst, rowLast, colFirst, colLast);
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Removes merged cells. Returns false if error occurs
tResult NVObjWorksheet::methodDelMerge( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: Row
	qlong omRow;
	if( getParamLong(pThreadData, 1, omRow) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, row, is required.";
		return ERR_BAD_PARAMS;
	}
	int row = static_cast<int> (omRow) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Column
	EXTfldval colVal;
	if ( getParamVar(pThreadData, 2, colVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, col, is required.";
		return ERR_BAD_PARAMS;
	} 
	int col = NVObjWorkbook::getColNumber(colVal) - 1; // Index off 0 for LibXL's sake
	if(col < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter, col.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	bool retBool = sheet->delMerge(row, col);
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Sets/removes a horizontal page break. Returns false if error occurs.
tResult NVObjWorksheet::methodSetHorzPageBreak( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: Row
	qlong omRow;
	if( getParamLong(pThreadData, 1, omRow) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, row, is required.";
		return ERR_BAD_PARAMS;
	}
	int row = static_cast<int> (omRow) - 1; // Index off 0 for LibXL's sake
	
	qbool omPageBreak;
	bool retBool;
	if ( getParamBool(pThreadData, 2, omPageBreak) == qtrue ) {
		// Call LibXL with page break param
		retBool = sheet->setHorPageBreak(row, getBoolFromQBool(omPageBreak));
	} else {
		// Call LibXL without page break param
		retBool = sheet->setHorPageBreak(row);
	}
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Sets/removes a vertical page break. Returns false if error occurs.
tResult NVObjWorksheet::methodSetVertPageBreak( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: Column
	EXTfldval colVal;
	if ( getParamVar(pThreadData, 1, colVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, col, is required.";
		return ERR_BAD_PARAMS;
	} 
	int col = NVObjWorkbook::getColNumber(colVal) - 1; // Index off 0 for LibXL's sake
	if(col < 0) {
		pThreadData->mExtraErrorText = "Unknown first parameter, col.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	qbool omPageBreak;
	bool retBool;
	if ( getParamBool(pThreadData, 2, omPageBreak) == qtrue ) {
		// Call LibXL with page break param
		retBool = sheet->setVerPageBreak(col, getBoolFromQBool(omPageBreak));
	} else {
		// Call LibXL without page break param
		retBool = sheet->setVerPageBreak(col);
	}
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Splits a sheet at position (row, col).
tResult NVObjWorksheet::methodSplit( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: Row
	qlong omRow;
	if( getParamLong(pThreadData, 1, omRow) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, row, is required.";
		return ERR_BAD_PARAMS;
	}
	int row = static_cast<int> (omRow) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Column
	EXTfldval colVal;
	if ( getParamVar(pThreadData, 2, colVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, col, is required.";
		return ERR_BAD_PARAMS;
	} 
	int col = NVObjWorkbook::getColNumber(colVal) - 1; // Index off 0 for LibXL's sake
	if(col < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter, col.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	sheet->split(row, col);
	
	return METHOD_DONE_RETURN;
}

// Groups rows from rowFirst to rowLast. Returns false if error occurs.
tResult NVObjWorksheet::methodGroupRows( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: First row
	qlong omFirstRow;
	if( getParamLong(pThreadData, 1, omFirstRow) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, rowFirst, is required.";
		return ERR_BAD_PARAMS;
	}
	int firstRow = static_cast<int> (omFirstRow) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Last row
	qlong omLastRow;
	if( getParamLong(pThreadData, 2, omLastRow) != qtrue) {
		pThreadData->mExtraErrorText = "Second parameter, rowLast, is required.";
		return ERR_BAD_PARAMS;
	}
	int lastRow = static_cast<int> (omLastRow) - 1; // Index off 0 for LibXL's sake
	
	qbool omCollapsed;
	bool retBool;
	if ( getParamBool(pThreadData, 2, omCollapsed) == qtrue ) {
		// Call LibXL with page break param
		retBool = sheet->groupRows(firstRow, lastRow, getBoolFromQBool(omCollapsed));
	} else {
		// Call LibXL without page break param
		retBool = sheet->groupRows(firstRow, lastRow);
	}
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Groups columns from colFirst to colLast. Returns false if error occurs.
tResult NVObjWorksheet::methodGroupCols( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Param 1: First col
	EXTfldval firstColVal;
	if ( getParamVar(pThreadData, 1, firstColVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, firstCol, is required.";
		return ERR_BAD_PARAMS;
	} 
	int firstCol = NVObjWorkbook::getColNumber(firstColVal) - 1; // Index off 0 for LibXL's sake
	if(firstCol < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter, firstCol.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 2: Second col
	EXTfldval lastColVal;
	if ( getParamVar(pThreadData, 2, lastColVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, lastCol, is required.";
		return ERR_BAD_PARAMS;
	} 
	int lastCol = NVObjWorkbook::getColNumber(lastColVal) - 1; // Index off 0 for LibXL's sake
	if(lastCol < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter, lastCol.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	qbool omCollapsed;
	bool retBool;
	if ( getParamBool(pThreadData, 3, omCollapsed) == qtrue ) {
		// Call LibXL with page break param
		retBool = sheet->groupCols(firstCol, lastCol, getBoolFromQBool(omCollapsed));
	} else {
		// Call LibXL without page break param
		retBool = sheet->groupCols(firstCol, lastCol);
	}
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Clears cells in specified area.
tResult NVObjWorksheet::methodClear( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	qlong omRowFirst, omRowLast;
	EXTfldval colFirstVal, colLastVal; 
	int rowFirst = 0, rowLast = MAX_ROWS, colFirst = 0, colLast = MAX_COLS;
	
	// Param 1: (Optional) First row to clear
	if( getParamLong(pThreadData, 1, omRowFirst) == qtrue ) {
		rowFirst = static_cast<int> (omRowFirst) - 1; // Index off 0 for LibXL's sake
	}
	
	// Param 2: (Optional) Last row to clear
	if( getParamLong(pThreadData, 2, omRowLast) == qtrue ) {
		rowLast = static_cast<int> (omRowLast) - 1; // Index off 0 for LibXL's sake
	}
	
	// Param 3: (Optional) First col to clear
	if ( getParamVar(pThreadData, 3, colFirstVal) == qtrue ) {
		colFirst = NVObjWorkbook::getColNumber(colFirstVal) - 1; // Index off 0 for LibXL's sake
		if(colFirst < 0) {
			pThreadData->mExtraErrorText = "Unknown third parameter, colFirst.  Acceptable types are Integer and Character.";
			return ERR_BAD_PARAMS;
		}
	} 
	
	// Param 4: (Optional) Last col to clear
	if ( getParamVar(pThreadData, 4, colLastVal) == qtrue ) {
		colLast = NVObjWorkbook::getColNumber(colLastVal) - 1; // Index off 0 for LibXL's sake
		if(colLast < 0) {
			pThreadData->mExtraErrorText = "Unknown fourth parameter, colLast.  Acceptable types are Integer and Character.";
			return ERR_BAD_PARAMS;
		}
	} 
	
	// Call LibXL
	sheet->clear(rowFirst, rowLast, colFirst, colLast);
	
	return METHOD_DONE_RETURN;
}

// Inserts rows from rowFirst to rowLast. Returns false if error occurs.
tResult NVObjWorksheet::methodInsertRow( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	qlong omRowFirst, omRowLast; 
	int rowFirst, rowLast;
	
	// Param 1: First row to insert
	if( getParamLong(pThreadData, 1, omRowFirst) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, rowFirst, is required.";
		return ERR_BAD_PARAMS;
	}
	rowFirst = static_cast<int> (omRowFirst) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Last row to insert
	if( getParamLong(pThreadData, 2, omRowLast) != qtrue) {
		pThreadData->mExtraErrorText = "Second parameter, rowLast, is required.";
		return ERR_BAD_PARAMS;
	}
	rowLast = static_cast<int> (omRowLast) - 1; // Index off 0 for LibXL's sake
	
	// Call LibXL
	bool retBool = sheet->insertRow(rowFirst, rowLast);
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Inserts cols from colFirst to colLast. Returns false if error occurs.
tResult NVObjWorksheet::methodInsertCol( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	EXTfldval colFirstVal, colLastVal;
	int colFirst, colLast;
	
	// Param 1: First col to insert
	if ( getParamVar(pThreadData, 1, colFirstVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, colFirst, is required.";
		return ERR_BAD_PARAMS;
	} 
	colFirst = NVObjWorkbook::getColNumber(colFirstVal) - 1; // Index off 0 for LibXL's sake
	if(colFirst < 0) {
		pThreadData->mExtraErrorText = "Unknown first parameter, colFirst.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 2: Last col to insert
	if ( getParamVar(pThreadData, 2, colLastVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, colLast, is required.";
		return ERR_BAD_PARAMS;
	} 
	colLast = NVObjWorkbook::getColNumber(colLastVal) - 1; // Index off 0 for LibXL's sake
	if(colLast < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter, colLast.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	bool retBool = sheet->insertCol(colFirst, colLast);
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Removes rows from rowFirst to rowLast. Returns false if error occurs.
tResult NVObjWorksheet::methodRemoveRow( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	qlong omRowFirst, omRowLast; 
	int rowFirst, rowLast;
	
	// Param 1: First row to remove
	if( getParamLong(pThreadData, 1, omRowFirst) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, rowFirst, is required.";
		return ERR_BAD_PARAMS;
	}
	rowFirst = static_cast<int> (omRowFirst) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Last row to remove
	if( getParamLong(pThreadData, 2, omRowLast) != qtrue) {
		pThreadData->mExtraErrorText = "Second parameter, rowLast, is required.";
		return ERR_BAD_PARAMS;
	}
	rowLast = static_cast<int> (omRowLast) - 1; // Index off 0 for LibXL's sake
	
	// Call LibXL
	bool retBool = sheet->removeRow(rowFirst, rowLast);
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Removes cols from colFirst to colLast. Returns false if error occurs.
tResult NVObjWorksheet::methodRemoveCol( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	EXTfldval colFirstVal, colLastVal;
	int colFirst, colLast;
	
	// Param 1: First col to remove
	if ( getParamVar(pThreadData, 1, colFirstVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, colFirst, is required.";
		return ERR_BAD_PARAMS;
	} 
	colFirst = NVObjWorkbook::getColNumber(colFirstVal) - 1; // Index off 0 for LibXL's sake
	if(colFirst < 0) {
		pThreadData->mExtraErrorText = "Unknown first parameter, colFirst.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 2: Last col to remove
	if ( getParamVar(pThreadData, 2, colLastVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, colLast, is required.";
		return ERR_BAD_PARAMS;
	} 
	colLast = NVObjWorkbook::getColNumber(colLastVal) - 1; // Index off 0 for LibXL's sake
	if(colLast < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter, colLast.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	bool retBool = sheet->removeCol(colFirst, colLast);
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Sets repeated rows on each page from rowFirst to rowLast.
tResult NVObjWorksheet::methodSetPrintRepeatRows( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	qlong omRowFirst, omRowLast; 
	int rowFirst, rowLast;
	
	// Param 1: First row to repeat
	if( getParamLong(pThreadData, 1, omRowFirst) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, rowFirst, is required.";
		return ERR_BAD_PARAMS;
	}
	rowFirst = static_cast<int> (omRowFirst) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Last row to repeat
	if( getParamLong(pThreadData, 2, omRowLast) != qtrue) {
		pThreadData->mExtraErrorText = "Second parameter, rowLast, is required.";
		return ERR_BAD_PARAMS;
	}
	rowLast = static_cast<int> (omRowLast) - 1; // Index off 0 for LibXL's sake
	
	// Call LibXL
	sheet->setPrintRepeatRows(rowFirst, rowLast);
	
	return METHOD_DONE_RETURN;
}

// Sets repeated columns on each page from colFirst to colLast.
tResult NVObjWorksheet::methodSetPrintRepeatCols( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	EXTfldval colFirstVal, colLastVal;
	int colFirst, colLast;
	
	// Param 1: First col to repeat
	if ( getParamVar(pThreadData, 1, colFirstVal) != qtrue ) {
		pThreadData->mExtraErrorText = "First parameter, colFirst, is required.";
		return ERR_BAD_PARAMS;
	} 
	colFirst = NVObjWorkbook::getColNumber(colFirstVal) - 1; // Index off 0 for LibXL's sake
	if(colFirst < 0) {
		pThreadData->mExtraErrorText = "Unknown first parameter, colFirst.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 2: Last col to repeat
	if ( getParamVar(pThreadData, 2, colLastVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Second parameter, colLast, is required.";
		return ERR_BAD_PARAMS;
	} 
	colLast = NVObjWorkbook::getColNumber(colLastVal) - 1; // Index off 0 for LibXL's sake
	if(colLast < 0) {
		pThreadData->mExtraErrorText = "Unknown second parameter, colLast.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	sheet->setPrintRepeatCols(colFirst, colLast);
	
	return METHOD_DONE_RETURN;
}

// Sets the print area
tResult NVObjWorksheet::methodSetPrintArea( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	qlong omRowFirst, omRowLast;
	EXTfldval colFirstVal, colLastVal; 
	int rowFirst, rowLast, colFirst, colLast;
	
	// Param 1: First row for print area
	if( getParamLong(pThreadData, 1, omRowFirst) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, rowFirst, is required.";
		return ERR_BAD_PARAMS;
	}
	rowFirst = static_cast<int> (omRowFirst) - 1; // Index off 0 for LibXL's sake
	
	// Param 2: Last row for print area
	if( getParamLong(pThreadData, 2, omRowLast) != qtrue) {
		pThreadData->mExtraErrorText = "Second parameter, rowLast, is required.";
		return ERR_BAD_PARAMS;
	}
	rowLast = static_cast<int> (omRowLast) - 1; // Index off 0 for LibXL's sake
	
	// Param 3: First col for print area
	if ( getParamVar(pThreadData, 3, colFirstVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Third parameter, colFirst, is required.";
		return ERR_BAD_PARAMS;
	} 
	colFirst = NVObjWorkbook::getColNumber(colFirstVal) - 1; // Index off 0 for LibXL's sake
	if(colFirst < 0) {
		pThreadData->mExtraErrorText = "Unknown third parameter, colFirst.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 4: Last col for print area
	if ( getParamVar(pThreadData, 4, colLastVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Fourth parameter, colLast, is required.";
		return ERR_BAD_PARAMS;
	} 
	colLast = NVObjWorkbook::getColNumber(colLastVal) - 1; // Index off 0 for LibXL's sake
	if(colLast < 0) {
		pThreadData->mExtraErrorText = "Unknown fourth parameter, colLast.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	sheet->setPrintArea(rowFirst, rowLast, colFirst, colLast);
	
	return METHOD_DONE_RETURN;
}

// Clears repeated rows and columns on each page.
tResult NVObjWorksheet::methodClearPrintRepeats( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Call LibXL
	sheet->clearPrintRepeats();
	
	return METHOD_DONE_RETURN;
}

// Clears the print area.
tResult NVObjWorksheet::methodClearPrintArea( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Call LibXL
	sheet->clearPrintArea();
	
	return METHOD_DONE_RETURN;
}

// Sets the named range. Returns false if error occurs. 
tResult NVObjWorksheet::methodSetNamedRange( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	qlong omRowFirst, omRowLast;
	EXTfldval nameVal, colFirstVal, colLastVal; 
	int rowFirst, rowLast, colFirst, colLast;
	std::wstring nameString;
	
	// Param 1: First row for print area
	if( getParamVar(pThreadData, 1, nameVal) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, name, is required.";
		return ERR_BAD_PARAMS;
	}
	nameString = getWStringFromEXTFldVal(nameVal);
	if (nameString.empty()) {
		pThreadData->mExtraErrorText = "Invalid First parameter. Name cannot be blank.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 2: First row for print area
	if( getParamLong(pThreadData, 2, omRowFirst) != qtrue) {
		pThreadData->mExtraErrorText = "Second parameter, rowFirst, is required.";
		return ERR_BAD_PARAMS;
	}
	rowFirst = static_cast<int> (omRowFirst) - 1; // Index off 0 for LibXL's sake
	
	// Param 3: Last row for print area
	if( getParamLong(pThreadData, 3, omRowLast) != qtrue) {
		pThreadData->mExtraErrorText = "Third parameter, rowLast, is required.";
		return ERR_BAD_PARAMS;
	}
	rowLast = static_cast<int> (omRowLast) - 1; // Index off 0 for LibXL's sake
	
	// Param 4: First col for print area
	if ( getParamVar(pThreadData, 4, colFirstVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Fourth parameter, colFirst, is required.";
		return ERR_BAD_PARAMS;
	} 
	colFirst = NVObjWorkbook::getColNumber(colFirstVal) - 1; // Index off 0 for LibXL's sake
	if(colFirst < 0) {
		pThreadData->mExtraErrorText = "Unknown fourth parameter, colFirst.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Param 5: Last col for print area
	if ( getParamVar(pThreadData, 5, colLastVal) != qtrue ) {
		pThreadData->mExtraErrorText = "Fifth parameter, colLast, is required.";
		return ERR_BAD_PARAMS;
	} 
	colLast = NVObjWorkbook::getColNumber(colLastVal) - 1; // Index off 0 for LibXL's sake
	if(colLast < 0) {
		pThreadData->mExtraErrorText = "Unknown fifth parameter, colLast.  Acceptable types are Integer and Character.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	bool retBool = sheet->setNamedRange(nameString.c_str(), rowFirst, rowLast, colFirst, colLast);
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

// Deletes the named range by name. Returns false if error occurs.
tResult NVObjWorksheet::methodDelNamedRange( tThreadData* pThreadData, qshort pParamCount ) {
	
	// Define all parameters
	EXTfldval nameVal; 
	std::wstring nameString;
	
	// Param 1: First row for print area
	if( getParamVar(pThreadData, 1, nameVal) != qtrue) {
		pThreadData->mExtraErrorText = "First parameter, name, is required.";
		return ERR_BAD_PARAMS;
	}
	nameString = getWStringFromEXTFldVal(nameVal);
	if (nameString.empty()) {
		pThreadData->mExtraErrorText = "Invalid First parameter. Name cannot be blank.";
		return ERR_BAD_PARAMS;
	}
	
	// Call LibXL
	bool retBool = sheet->delNamedRange(nameString.c_str());
	
	// Return value to Omnis
	EXTfldval retVal;
	getEXTFldValFromBool(retVal, retBool);
	ECOaddParam(pThreadData->mEci, &retVal);
	
	return METHOD_DONE_RETURN;
}

/**************************************************************************************************
 **                              INTERNAL      METHODS                                           **
 **************************************************************************************************/

bool NVObjWorksheet::checkDependencies() {
	if (numFormats && book && sheet)
		return true;
	else
		return false;
}

void NVObjWorksheet::setDependencies(shared_ptr<NumFormatMgr> nf, shared_ptr<Book> b, Sheet* s) {
	setNumFormatMgr(nf);
	setBook(b);
	setSheet(s);
}

shared_ptr<NumFormatMgr> NVObjWorksheet::getNumFormatMgr() {
	return numFormats;
}

shared_ptr<Book> NVObjWorksheet::getBook() {
	return book;
}

Sheet* NVObjWorksheet::getSheet() {
	return sheet;
}

void NVObjWorksheet::setNumFormatMgr(shared_ptr<NumFormatMgr> nf) {
	numFormats = nf;
}

void NVObjWorksheet::setBook(shared_ptr<Book> b) {
	book = b;
}

void NVObjWorksheet::setSheet(Sheet* s) {
	sheet = s;
}



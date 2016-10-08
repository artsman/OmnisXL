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

#include "Constants.he"

using namespace libxl;
using std::map;

// Build the maps used by getColorConstant and getOmnisColorConstant
static map<int, Color> colorLookup;
static map<Color, int> omnisColorLookup;
void LibXLConstants::buildColorLookup() {
	if (colorLookup.size() > 0) // Pseudo-singleton for color lookup static variables.
		return;
	
	// Build main lookup (based on constant values defined in ExcelFormat.rc)
 	colorLookup[1]  = COLOR_BLACK;	
 	colorLookup[2]  = COLOR_WHITE;	
 	colorLookup[3]  = COLOR_RED;	
 	colorLookup[4]  = COLOR_BRIGHTGREEN;	
 	colorLookup[5]  = COLOR_BLUE;	
 	colorLookup[6]  = COLOR_YELLOW;	
 	colorLookup[7]  = COLOR_PINK;	
 	colorLookup[8]  = COLOR_TURQUOISE;	
 	colorLookup[9]  = COLOR_DARKRED;	
 	colorLookup[10] = COLOR_GREEN;	
 	colorLookup[11] = COLOR_DARKBLUE;	
 	colorLookup[12] = COLOR_DARKYELLOW;	
 	colorLookup[13] = COLOR_VIOLET;	
 	colorLookup[14] = COLOR_TEAL;	
 	colorLookup[15] = COLOR_GRAY25;	
 	colorLookup[16] = COLOR_GRAY50;	
 	colorLookup[17] = COLOR_PERIWINKLE_CF;	
 	colorLookup[18] = COLOR_PLUM_CF;	
 	colorLookup[19] = COLOR_IVORY_CF;	
 	colorLookup[20] = COLOR_LIGHTTURQUOISE_CF;	
 	colorLookup[21] = COLOR_DARKPURPLE_CF;	
 	colorLookup[22] = COLOR_CORAL_CF;	
 	colorLookup[23] = COLOR_OCEANBLUE_CF;	
 	colorLookup[24] = COLOR_ICEBLUE_CF;	
 	colorLookup[25] = COLOR_DARKBLUE_CL;	
 	colorLookup[26] = COLOR_PINK_CL;	
 	colorLookup[27] = COLOR_YELLOW_CL;	
 	colorLookup[28] = COLOR_TURQUOISE_CL;	
 	colorLookup[29] = COLOR_VIOLET_CL;	
 	colorLookup[30] = COLOR_DARKRED_CL;	
 	colorLookup[31] = COLOR_TEAL_CL;	
 	colorLookup[32] = COLOR_BLUE_CL;	
 	colorLookup[33] = COLOR_SKYBLUE;	
 	colorLookup[34] = COLOR_LIGHTTURQUOISE;	
 	colorLookup[35] = COLOR_LIGHTGREEN;	
 	colorLookup[36] = COLOR_LIGHTYELLOW;	
 	colorLookup[37] = COLOR_PALEBLUE;	
 	colorLookup[38] = COLOR_ROSE;	
 	colorLookup[39] = COLOR_LAVENDER;	
 	colorLookup[40] = COLOR_TAN;	
 	colorLookup[41] = COLOR_LIGHTBLUE;	
 	colorLookup[42] = COLOR_AQUA;	
 	colorLookup[43] = COLOR_LIME;	
 	colorLookup[44] = COLOR_GOLD;	
 	colorLookup[45] = COLOR_LIGHTORANGE;	
 	colorLookup[46] = COLOR_ORANGE;	
 	colorLookup[47] = COLOR_BLUEGRAY;	
 	colorLookup[48] = COLOR_GRAY40;	
 	colorLookup[49] = COLOR_DARKTEAL;	
 	colorLookup[50] = COLOR_SEAGREEN;	
 	colorLookup[51] = COLOR_DARKGREEN;	
 	colorLookup[52] = COLOR_OLIVEGREEN;	
 	colorLookup[53] = COLOR_BROWN;	
 	colorLookup[54] = COLOR_PLUM;	
 	colorLookup[55] = COLOR_INDIGO;	
 	colorLookup[56] = COLOR_GRAY80;	
 	colorLookup[57] = COLOR_DEFAULT_FOREGROUND;	
 	colorLookup[58] = COLOR_DEFAULT_BACKGROUND;
	colorLookup[59] = COLOR_TOOLTIP;
	colorLookup[60] = COLOR_AUTO;
	
	// Build reverse lookup
	map<int, Color>::iterator it;
	for ( it=colorLookup.begin() ; it != colorLookup.end(); it++ ) {
		omnisColorLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidColor(int omnisColorConstant) {
	buildColorLookup();
	
	map<int,Color>::iterator it;
	it = colorLookup.find(omnisColorConstant);
	if (it != colorLookup.end())
		return true;
	else
		return false;
}

// Get the LibXL color constant for the Omnis color constant
Color LibXLConstants::getColorConstant(int omnisColorConstant) {
	buildColorLookup();
	
	map<int,Color>::iterator it;
	it = colorLookup.find(omnisColorConstant);
	if (it == colorLookup.end())
		return COLOR_AUTO;
	else
		return it->second;;
}

// Get the Omnis color constant for the LibXL color constant
int LibXLConstants::getOmnisColorConstant(Color colorConstant) {
	buildColorLookup();
	
	map<Color,int>::iterator it;
	it = omnisColorLookup.find(colorConstant);
	if (it == omnisColorLookup.end())
		return 0;
	else
		return it->second;;
}

// Number Format
static map<int, int> numFormatLookup;
static map<int, int> omnisNumFormatLookup;
void LibXLConstants::buildNumFormatLookup() {
	if (numFormatLookup.size() > 0) // Pseudo-singleton for number format lookup static variables.
		return;
	
	numFormatLookup[1]  = NUMFORMAT_GENERAL;
	numFormatLookup[2]  = NUMFORMAT_NUMBER;
	numFormatLookup[3]  = NUMFORMAT_NUMBER_D2;
	numFormatLookup[4]  = NUMFORMAT_NUMBER_SEP;
	numFormatLookup[5]  = NUMFORMAT_NUMBER_SEP_D2;
	numFormatLookup[6]  = NUMFORMAT_CURRENCY_NEGBRA;
	numFormatLookup[7]  = NUMFORMAT_CURRENCY_NEGBRARED;
	numFormatLookup[8]  = NUMFORMAT_CURRENCY_D2_NEGBRA;
	numFormatLookup[9]  = NUMFORMAT_CURRENCY_D2_NEGBRARED;
	numFormatLookup[10] = NUMFORMAT_PERCENT;
	numFormatLookup[11] = NUMFORMAT_PERCENT_D2;
	numFormatLookup[12] = NUMFORMAT_SCIENTIFIC_D2;
	numFormatLookup[13] = NUMFORMAT_FRACTION_ONEDIG;
	numFormatLookup[14] = NUMFORMAT_FRACTION_TWODIG;
	numFormatLookup[15] = NUMFORMAT_DATE;
	numFormatLookup[16] = NUMFORMAT_CUSTOM_D_MON_YY;
	numFormatLookup[17] = NUMFORMAT_CUSTOM_D_MON;
	numFormatLookup[18] = NUMFORMAT_CUSTOM_MON_YY;
	numFormatLookup[19] = NUMFORMAT_CUSTOM_HMM_AM;
	numFormatLookup[20] = NUMFORMAT_CUSTOM_HMMSS_AM;
	numFormatLookup[21] = NUMFORMAT_CUSTOM_HMM;
	numFormatLookup[22] = NUMFORMAT_CUSTOM_HMMSS;
	numFormatLookup[23] = NUMFORMAT_CUSTOM_MDYYYY_HMM;
	numFormatLookup[24] = NUMFORMAT_NUMBER_SEP_NEGBRA;
	numFormatLookup[25] = NUMFORMAT_NUMBER_SEP_NEGBRARED;
	numFormatLookup[26] = NUMFORMAT_NUMBER_D2_SEP_NEGBRA;
	numFormatLookup[27] = NUMFORMAT_NUMBER_D2_SEP_NEGBRARED;
	numFormatLookup[28] = NUMFORMAT_ACCOUNT;
	numFormatLookup[29] = NUMFORMAT_ACCOUNTCUR;
	numFormatLookup[30] = NUMFORMAT_ACCOUNT_D2;
	numFormatLookup[31] = NUMFORMAT_ACCOUNT_D2_CUR;
	numFormatLookup[32] = NUMFORMAT_CUSTOM_MMSS;
	numFormatLookup[33] = NUMFORMAT_CUSTOM_H0MMSS;
	numFormatLookup[34] = NUMFORMAT_CUSTOM_MMSS0;
	numFormatLookup[35] = NUMFORMAT_CUSTOM_000P0E_PLUS0;
	numFormatLookup[36] = NUMFORMAT_TEXT;
	
	// Build reverse lookup
	map<int,int>::iterator it;
	for ( it=numFormatLookup.begin() ; it != numFormatLookup.end(); it++ ) {
		omnisNumFormatLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidNumFormat(int omnisNumFormatConstant) {
	buildNumFormatLookup();
	
	map<int,int>::iterator it;
	it = numFormatLookup.find(omnisNumFormatConstant);
	if (it != numFormatLookup.end())
		return true;
	else
		return false;
}

int LibXLConstants::getNumFormatConstant(int omnisNumFormatConstant) {
	buildNumFormatLookup();
	
	map<int,int>::iterator it;
	it = numFormatLookup.find(omnisNumFormatConstant);
	if (it == numFormatLookup.end())
		return 0;
	else
		return it->second;;
}

int LibXLConstants::getOmnisNumFormatConstant(int numFormatConstant) {
	buildNumFormatLookup();
	
	map<int,int>::iterator it;
	it = omnisNumFormatLookup.find(numFormatConstant);
	if (it == omnisNumFormatLookup.end())
		return 0;
	else
		return it->second;;
}

// Horizontal Align
static map<int, AlignH> alignHLookup;
static map<AlignH, int> omnisAlignHLookup;
void LibXLConstants::buildAlignHLookup() {
	if (alignHLookup.size() > 0) // Pseudo-singleton for horizontal alignment lookup static variables.
		return;
	
	alignHLookup[1] = ALIGNH_GENERAL;
	alignHLookup[2] = ALIGNH_LEFT;
	alignHLookup[3] = ALIGNH_CENTER;
	alignHLookup[4] = ALIGNH_RIGHT;
	alignHLookup[5] = ALIGNH_FILL;
	alignHLookup[6] = ALIGNH_JUSTIFY;
	alignHLookup[7] = ALIGNH_MERGE;
	alignHLookup[8] = ALIGNH_DISTRIBUTED;
	
	// Build reverse lookup
	map<int,AlignH>::iterator it;
	for ( it=alignHLookup.begin() ; it != alignHLookup.end(); it++ ) {
		omnisAlignHLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidAlignH(int omnisAlignHConstant) {
	buildAlignHLookup();
	
	map<int,AlignH>::iterator it;
	it = alignHLookup.find(omnisAlignHConstant);
	if (it != alignHLookup.end())
		return true;
	else
		return false;
}

AlignH LibXLConstants::getAlignHConstant(int omnisAlignHConstant) {
	buildAlignHLookup();
	
	map<int,AlignH>::iterator it;
	it = alignHLookup.find(omnisAlignHConstant);
	if (it == alignHLookup.end())
		return ALIGNH_GENERAL;
	else
		return it->second;;
}

int LibXLConstants::getOmnisAlignHConstant(AlignH alignHConstant) {
	buildAlignHLookup();
	
	map<AlignH, int>::iterator it;
	it = omnisAlignHLookup.find(alignHConstant);
	if (it == omnisAlignHLookup.end())
		return 0;
	else
		return it->second;;
}


// Vertical Align
static map<int, AlignV> alignVLookup;
static map<AlignV, int> omnisAlignVLookup;
void LibXLConstants::buildAlignVLookup() {
	if (alignVLookup.size() > 0) // Pseudo-singleton for vertical alignment lookup static variables.
		return;
	
	alignVLookup[1] = ALIGNV_TOP;
	alignVLookup[2] = ALIGNV_CENTER;
	alignVLookup[3] = ALIGNV_BOTTOM;
	alignVLookup[4] = ALIGNV_JUSTIFY;
	alignVLookup[5] = ALIGNV_DISTRIBUTED;
	
	// Build reverse lookup
	map<int,AlignV>::iterator it;
	for ( it=alignVLookup.begin() ; it != alignVLookup.end(); it++ ) {
		omnisAlignVLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidAlignV(int omnisAlignVConstant) {
	buildAlignVLookup();
	
	map<int,AlignV>::iterator it;
	it = alignVLookup.find(omnisAlignVConstant);
	if (it != alignVLookup.end())
		return true;
	else
		return false;
}

AlignV LibXLConstants::getAlignVConstant(int omnisAlignVConstant) {
	buildAlignVLookup();
	
	map<int,AlignV>::iterator it;
	it = alignVLookup.find(omnisAlignVConstant);
	if (it == alignVLookup.end())
		return ALIGNV_TOP;
	else
		return it->second;
}

int LibXLConstants::getOmnisAlignVConstant(AlignV alignVConstant) {
	buildAlignVLookup();
	
	map<AlignV, int>::iterator it;
	it = omnisAlignVLookup.find(alignVConstant);
	if (it == omnisAlignVLookup.end())
		return 0;
	else
		return it->second;
}


// Border Style
static map<int, BorderStyle> borderStyleLookup;
static map<BorderStyle, int> omnisBorderStyleLookup;
void LibXLConstants::buildBorderStyleLookup() {
	if (borderStyleLookup.size() > 0) // Pseudo-singleton for border style lookup static variables.
		return;
	
	borderStyleLookup[1]  = BORDERSTYLE_NONE;
	borderStyleLookup[2]  = BORDERSTYLE_THIN;
	borderStyleLookup[3]  = BORDERSTYLE_MEDIUM;
	borderStyleLookup[4]  = BORDERSTYLE_DASHED;
	borderStyleLookup[5]  = BORDERSTYLE_DOTTED;
	borderStyleLookup[6]  = BORDERSTYLE_THICK;
	borderStyleLookup[7]  = BORDERSTYLE_DOUBLE;
	borderStyleLookup[8]  = BORDERSTYLE_HAIR;
	borderStyleLookup[9]  = BORDERSTYLE_MEDIUMDASHED;
	borderStyleLookup[10] = BORDERSTYLE_DASHDOT;
	borderStyleLookup[11] = BORDERSTYLE_MEDIUMDASHDOT;
	borderStyleLookup[12] = BORDERSTYLE_DASHDOTDOT;
	borderStyleLookup[13] = BORDERSTYLE_MEDIUMDASHDOTDOT;
	borderStyleLookup[14] = BORDERSTYLE_SLANTDASHDOT;
	
	// Build reverse lookup
	map<int,BorderStyle>::iterator it;
	for ( it=borderStyleLookup.begin() ; it != borderStyleLookup.end(); it++ ) {
		omnisBorderStyleLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidBorder(int omnisBorderStyleConstant) {
	buildBorderStyleLookup();
	
	map<int,BorderStyle>::iterator it;
	it = borderStyleLookup.find(omnisBorderStyleConstant);
	if (it != borderStyleLookup.end())
		return true;
	else
		return false;
}

BorderStyle LibXLConstants::getBorderStyleConstant(int omnisBorderStyleConstant) {
	buildBorderStyleLookup();
	
	map<int, BorderStyle>::iterator it;
	it = borderStyleLookup.find(omnisBorderStyleConstant);
	if (it == borderStyleLookup.end())
		return BORDERSTYLE_NONE;
	else
		return it->second;
}

int LibXLConstants::getOmnisBorderStyleConstant(BorderStyle borderStyleConstant) {
	buildBorderStyleLookup();
	
	map<BorderStyle, int>::iterator it;
	it = omnisBorderStyleLookup.find(borderStyleConstant);
	if (it == omnisBorderStyleLookup.end())
		return 0;
	else
		return it->second;
}


// Diagonal Border Style
static map<int, BorderDiagonal> borderDiagonalLookup;
static map<BorderDiagonal, int> omnisBorderDiagonalLookup;
void LibXLConstants::buildBorderDiagonalLookup() {
	if (borderDiagonalLookup.size() > 0) // Pseudo-singleton for border style lookup static variables.
		return;
	
	borderDiagonalLookup[1] = BORDERDIAGONAL_NONE;
	borderDiagonalLookup[2] = BORDERDIAGONAL_DOWN;
	borderDiagonalLookup[3] = BORDERDIAGONAL_UP;
	borderDiagonalLookup[4] = BORDERDIAGONAL_BOTH;
	
	// Build reverse lookup
	map<int,BorderDiagonal>::iterator it;
	for ( it=borderDiagonalLookup.begin() ; it != borderDiagonalLookup.end(); it++ ) {
		omnisBorderDiagonalLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidBorderDiagonal(int omnisBorderDiagonalConstant) {
	buildBorderDiagonalLookup();
	
	map<int,BorderDiagonal>::iterator it;
	it = borderDiagonalLookup.find(omnisBorderDiagonalConstant);
	if (it != borderDiagonalLookup.end())
		return true;
	else
		return false;
}

BorderDiagonal LibXLConstants::getBorderDiagonalConstant(int omnisBorderDiagonalConstant) {
	buildBorderDiagonalLookup();
	
	map<int, BorderDiagonal>::iterator it;
	it = borderDiagonalLookup.find(omnisBorderDiagonalConstant);
	if (it == borderDiagonalLookup.end())
		return BORDERDIAGONAL_NONE;
	else
		return it->second;
}

int LibXLConstants::getOmnisBorderDiagonalConstant(BorderDiagonal borderDiagonalConstant) {
	buildBorderDiagonalLookup();
	
	map<BorderDiagonal, int>::iterator it;
	it = omnisBorderDiagonalLookup.find(borderDiagonalConstant);
	if (it == omnisBorderDiagonalLookup.end())
		return 0;
	else
		return it->second;
}


// Fill Pattern Style
static map<int, FillPattern> fillPatternLookup;
static map<FillPattern, int> omnisFillPatternLookup;
void LibXLConstants::buildFillPatternLookup() {
	if (fillPatternLookup.size() > 0) // Pseudo-singleton for border style lookup static variables.
		return;
	
	fillPatternLookup[1]  = FILLPATTERN_NONE;
	fillPatternLookup[2]  = FILLPATTERN_SOLID;
	fillPatternLookup[3]  = FILLPATTERN_GRAY50;
	fillPatternLookup[4]  = FILLPATTERN_GRAY75;
	fillPatternLookup[5]  = FILLPATTERN_GRAY25;
	fillPatternLookup[6]  = FILLPATTERN_HORSTRIPE;
	fillPatternLookup[7]  = FILLPATTERN_VERSTRIPE;
	fillPatternLookup[8]  = FILLPATTERN_REVDIAGSTRIPE;
	fillPatternLookup[9]  = FILLPATTERN_DIAGSTRIPE;
	fillPatternLookup[10] = FILLPATTERN_DIAGCROSSHATCH;
	fillPatternLookup[11] = FILLPATTERN_THICKDIAGCROSSHATCH;
	fillPatternLookup[12] = FILLPATTERN_THINHORSTRIPE;
	fillPatternLookup[13] = FILLPATTERN_THINVERSTRIPE;
	fillPatternLookup[14] = FILLPATTERN_THINREVDIAGSTRIPE;
	fillPatternLookup[15] = FILLPATTERN_THINDIAGSTRIPE;
	fillPatternLookup[16] = FILLPATTERN_THINHORCROSSHATCH;
	fillPatternLookup[17] = FILLPATTERN_THINDIAGCROSSHATCH;
	fillPatternLookup[18] = FILLPATTERN_GRAY12P5;
	fillPatternLookup[19] = FILLPATTERN_GRAY6P25;
	
	// Build reverse lookup
	map<int,FillPattern>::iterator it;
	for ( it=fillPatternLookup.begin() ; it != fillPatternLookup.end(); it++ ) {
		omnisFillPatternLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidFillPattern(int omnisFillPatternConstant) {
	buildFillPatternLookup();
	
	map<int,FillPattern>::iterator it;
	it = fillPatternLookup.find(omnisFillPatternConstant);
	if (it != fillPatternLookup.end())
		return true;
	else
		return false;
}

FillPattern LibXLConstants::getFillPatternConstant(int omnisFillPatternConstant) {
	buildFillPatternLookup();
	
	map<int, FillPattern>::iterator it;
	it = fillPatternLookup.find(omnisFillPatternConstant);
	if (it == fillPatternLookup.end())
		return FILLPATTERN_NONE;
	else
		return it->second;
}

int LibXLConstants::getOmnisFillPatternConstant(FillPattern fillPatternConstant) {
	buildFillPatternLookup();
	
	map<FillPattern, int>::iterator it;
	it = omnisFillPatternLookup.find(fillPatternConstant);
	if (it == omnisFillPatternLookup.end())
		return 0;
	else
		return it->second;
}


// Script Style
static map<int, Script> scriptLookup;
static map<Script, int> omnisScriptLookup;
void LibXLConstants::buildScriptLookup() {
	if (scriptLookup.size() > 0) // Pseudo-singleton for border style lookup static variables.
		return;
	
	scriptLookup[1] = SCRIPT_NORMAL;
	scriptLookup[2] = SCRIPT_SUPER;
	scriptLookup[3] = SCRIPT_SUB;
	
	// Build reverse lookup
	map<int,Script>::iterator it;
	for ( it=scriptLookup.begin() ; it != scriptLookup.end(); it++ ) {
		omnisScriptLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidScript(int omnisScriptConstant) {
	buildScriptLookup();
	
	map<int,Script>::iterator it;
	it = scriptLookup.find(omnisScriptConstant);
	if (it != scriptLookup.end())
		return true;
	else
		return false;
}

Script LibXLConstants::getScriptConstant(int omnisScriptConstant) {
	buildScriptLookup();
	
	map<int, Script>::iterator it;
	it = scriptLookup.find(omnisScriptConstant);
	if (it == scriptLookup.end())
		return SCRIPT_NORMAL;
	else
		return it->second;
}

int LibXLConstants::getOmnisScriptConstant(Script scriptConstant) {
	buildScriptLookup();
	
	map<Script, int>::iterator it;
	it = omnisScriptLookup.find(scriptConstant);
	if (it == omnisScriptLookup.end())
		return 0;
	else
		return it->second;
}

// Underline Style
static map<int, Underline> underlineLookup;
static map<Underline, int> omnisUnderlineLookup;
void LibXLConstants::buildUnderlineLookup() {
	if (underlineLookup.size() > 0) // Pseudo-singleton for border style lookup static variables.
		return;
	
	underlineLookup[1] = UNDERLINE_NONE;	
	underlineLookup[2] = UNDERLINE_SINGLE;	
	underlineLookup[3] = UNDERLINE_DOUBLE;	
	underlineLookup[4] = UNDERLINE_SINGLEACC;	
	underlineLookup[5] = UNDERLINE_DOUBLEACC;	
	
	// Build reverse lookup
	map<int,Underline>::iterator it;
	for ( it=underlineLookup.begin() ; it != underlineLookup.end(); it++ ) {
		omnisUnderlineLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidUnderline(int omnisUnderlineConstant) {
	buildUnderlineLookup();
	
	map<int,Underline>::iterator it;
	it = underlineLookup.find(omnisUnderlineConstant);
	if (it != underlineLookup.end())
		return true;
	else
		return false;
}

Underline LibXLConstants::getUnderlineConstant(int omnisUnderlineConstant) {
	buildUnderlineLookup();
	
	map<int, Underline>::iterator it;
	it = underlineLookup.find(omnisUnderlineConstant);
	if (it == underlineLookup.end())
		return UNDERLINE_NONE;
	else
		return it->second;
}

int LibXLConstants::getOmnisUnderlineConstant(Underline underlineConstant) {
	buildUnderlineLookup();
	
	map<Underline, int>::iterator it;
	it = omnisUnderlineLookup.find(underlineConstant);
	if (it == omnisUnderlineLookup.end())
		return 0;
	else
		return it->second;
}

// Build the maps used by getPaperConstant and getOmnisPaperConstant
//    Allocate the static variables
static map<int, Paper> paperLookup;
static map<Paper, int> omnisPaperLookup;
//    One time build of lookup
void LibXLConstants::buildPaperLookup() {
	if (paperLookup.size() > 0) // Pseudo-singleton for paper lookup static variables.
		return;
	
	// Build main lookup (based on constant values defined in ExcelFormat.rc)
	paperLookup[1]  = PAPER_DEFAULT;
	paperLookup[2]  = PAPER_LETTER;
	paperLookup[3]  = PAPER_LETTERSMALL;
	paperLookup[4]  = PAPER_TABLOID;
	paperLookup[5]  = PAPER_LEDGER;
	paperLookup[6]  = PAPER_LEGAL;
	paperLookup[7]  = PAPER_STATEMENT;
	paperLookup[8]  = PAPER_EXECUTIVE;
	paperLookup[9]  = PAPER_A3;
	paperLookup[10] = PAPER_A4;
	paperLookup[11] = PAPER_A4SMALL;
	paperLookup[12] = PAPER_A5;
	paperLookup[13] = PAPER_B4;
	paperLookup[14] = PAPER_B5;
	paperLookup[15] = PAPER_FOLIO;
	paperLookup[16] = PAPER_QUATRO;
	paperLookup[17] = PAPER_10x14;
	paperLookup[18] = PAPER_10x17;
	paperLookup[19] = PAPER_NOTE;
	paperLookup[20] = PAPER_ENVELOPE_9;
	paperLookup[21] = PAPER_ENVELOPE_10;
	paperLookup[22] = PAPER_ENVELOPE_11;
	paperLookup[23] = PAPER_ENVELOPE_12;
	paperLookup[24] = PAPER_ENVELOPE_14;
	paperLookup[25] = PAPER_C_SIZE;
	paperLookup[26] = PAPER_D_SIZE;
	paperLookup[27] = PAPER_E_SIZE;
	paperLookup[28] = PAPER_ENVELOPE_DL;
	paperLookup[29] = PAPER_ENVELOPE_C5;
	paperLookup[30] = PAPER_ENVELOPE_C3;
	paperLookup[31] = PAPER_ENVELOPE_C4;
	paperLookup[32] = PAPER_ENVELOPE_C6;
	paperLookup[33] = PAPER_ENVELOPE_C65;
	paperLookup[34] = PAPER_ENVELOPE_B4;
	paperLookup[35] = PAPER_ENVELOPE_B5;
	paperLookup[36] = PAPER_ENVELOPE_B6;
	paperLookup[37] = PAPER_ENVELOPE;
	paperLookup[38] = PAPER_ENVELOPE_MONARCH;
	paperLookup[39] = PAPER_US_ENVELOPE;
	paperLookup[40] = PAPER_FANFOLD;
	paperLookup[41] = PAPER_GERMAN_STD_FANFOLD;
	paperLookup[42] = PAPER_GERMAN_LEGAL_FANFOLD;
	
	// Build reverse lookup
	map<int,Paper>::iterator it;
	for ( it=paperLookup.begin() ; it != paperLookup.end(); it++ ) {
		omnisPaperLookup[(*it).second] = (*it).first;
	}
	
	return;
}

// Check if the constant is valid
bool LibXLConstants::isValidPaper(int omnisPaperConstant) {
	buildPaperLookup();
	
	map<int,Paper>::iterator it;
	it = paperLookup.find(omnisPaperConstant);
	if (it != paperLookup.end())
		return true;
	else
		return false;
}

// Get the LibXL paper constant for the Omnis paper constant
Paper LibXLConstants::getPaperConstant(int omnisPaperConstant) {
	buildPaperLookup();
	
	map<int, Paper>::iterator it;
	it = paperLookup.find(omnisPaperConstant);
	if (it == paperLookup.end())
		return PAPER_DEFAULT;
	else
		return it->second;
}

// Get the Omnis paper constant for the LibXL paper constant
int LibXLConstants::getOmnisPaperConstant(Paper paperConstant) {
	buildPaperLookup();
	
	map<Paper, int>::iterator it;
	it = omnisPaperLookup.find(paperConstant);
	if (it == omnisPaperLookup.end())
		return 0;
	else
		return it->second;
}

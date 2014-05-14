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

#include "NumFormatMgr.he"

using namespace libxl;
using boost::shared_ptr;

NumFormatMgr::NumFormatMgr(boost::shared_ptr<libxl::Book> b) : book(b) {
	buildFormatLookup();
}

NumFormatMgr::~NumFormatMgr() { }

// Retrieve a single format.  If it doesn't exist then create it for the book and store the pointer.
Format* NumFormatMgr::getFormat(int fmtConst) {
	Format* retFormat = formatLookup[fmtConst];
	if (!retFormat) {
		// Format doesn't exist.  Create a new one
		retFormat = book->addFormat();
		if (retFormat)
            retFormat->setNumFormat(fmtConst);
		formatLookup[fmtConst] = retFormat;
	}
	
	return retFormat;
}

// Build lookup and set all Format pointers to 0;
void NumFormatMgr::buildFormatLookup() {
	
	formatLookup[NUMFORMAT_GENERAL] = 0;
	formatLookup[NUMFORMAT_NUMBER] = 0;
	formatLookup[NUMFORMAT_NUMBER_D2] = 0;
	formatLookup[NUMFORMAT_NUMBER_SEP] = 0;
	formatLookup[NUMFORMAT_NUMBER_SEP_D2] = 0;
	formatLookup[NUMFORMAT_CURRENCY_NEGBRA] = 0;
	formatLookup[NUMFORMAT_CURRENCY_NEGBRARED] = 0;
	formatLookup[NUMFORMAT_CURRENCY_D2_NEGBRA] = 0;
	formatLookup[NUMFORMAT_CURRENCY_D2_NEGBRARED] = 0;
	formatLookup[NUMFORMAT_PERCENT] = 0;
	formatLookup[NUMFORMAT_PERCENT_D2] = 0;
	formatLookup[NUMFORMAT_SCIENTIFIC_D2] = 0;
	formatLookup[NUMFORMAT_FRACTION_ONEDIG] = 0;
	formatLookup[NUMFORMAT_FRACTION_TWODIG] = 0;
	formatLookup[NUMFORMAT_DATE] = 0;
	formatLookup[NUMFORMAT_CUSTOM_D_MON_YY] = 0;
	formatLookup[NUMFORMAT_CUSTOM_D_MON] = 0;
	formatLookup[NUMFORMAT_CUSTOM_MON_YY] = 0;
	formatLookup[NUMFORMAT_CUSTOM_HMM_AM] = 0;
	formatLookup[NUMFORMAT_CUSTOM_HMMSS_AM] = 0;
	formatLookup[NUMFORMAT_CUSTOM_HMM] = 0;
	formatLookup[NUMFORMAT_CUSTOM_HMMSS] = 0;
	formatLookup[NUMFORMAT_CUSTOM_MDYYYY_HMM] = 0;
	formatLookup[NUMFORMAT_NUMBER_SEP_NEGBRA] = 0;
	formatLookup[NUMFORMAT_NUMBER_SEP_NEGBRARED] = 0;
	formatLookup[NUMFORMAT_NUMBER_D2_SEP_NEGBRA] = 0;
	formatLookup[NUMFORMAT_NUMBER_D2_SEP_NEGBRARED] = 0;
	formatLookup[NUMFORMAT_ACCOUNT] = 0;
	formatLookup[NUMFORMAT_ACCOUNTCUR] = 0;
	formatLookup[NUMFORMAT_ACCOUNT_D2] = 0;
	formatLookup[NUMFORMAT_ACCOUNT_D2_CUR] = 0;
	formatLookup[NUMFORMAT_CUSTOM_MMSS] = 0;
	formatLookup[NUMFORMAT_CUSTOM_H0MMSS] = 0;
	formatLookup[NUMFORMAT_CUSTOM_MMSS0] = 0;
	formatLookup[NUMFORMAT_CUSTOM_000P0E_PLUS0] = 0;
	formatLookup[NUMFORMAT_TEXT] = 0;
	
	return;
}
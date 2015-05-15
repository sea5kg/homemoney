//---------------------------------------------------------------------------

#ifndef helpersH
#define helpersH
#include <vcl.h>
//---------------------------------------------------------------------------

struct exlSheet {
	int Number;
	UnicodeString Name;
};

struct exlClass {
	UnicodeString Class;
	UnicodeString Name;
	UnicodeString Comment;
	UnicodeString Monthes;
};

struct exlMonth {
	UnicodeString Name;
	UnicodeString Class;
	double Price;
	UnicodeString Month;
//	std::vector<UnicodeString> Hyperlinks;
	UnicodeString LinkToClassification;
};

struct exlSumClass {
	UnicodeString Name;
    double Sum;
};

#endif

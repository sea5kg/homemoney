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

struct exlSumClass {
	UnicodeString Name;
    double Sum;
};

#endif

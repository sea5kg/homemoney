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
	double Price;
	UnicodeString Month;
};

#endif

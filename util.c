#pragma warning (disable: 4996) // _CRT_SECURE_NO_WARNINGS

#include "util.h"

// Compares a pascal string and a null-terminated C-string to see if they
// are equal. Case insensitive.
int lpwstricmp(LPWSTR s, LPWSTR t)
{
    int i;

    if (wcslen(s) != *t)
        return 1;

    for (i = 1; i <= s[0]; i++)
    {
        if (towlower(s[i - 1]) != towlower(t[i]))
            return 1;
    }

    return 0;
}

// Create counted Unicode wchar string from null-terminated ASCII input
wchar_t *new_xl12string(const char *text)
{
    size_t len = strlen(text);
    if (!text || !len)
        return NULL;
    if (len > 255)
        len = 255; // truncate
    wchar_t *p = (wchar_t *)malloc((len + 2) * sizeof(wchar_t));
    if (!p) return NULL;
    mbstowcs(p + 1, text, len);
    p[0] = (wchar_t)len; // string p[1] is NOT null terminated
    p[len + 1] = 0; // now it is
    return p;
}

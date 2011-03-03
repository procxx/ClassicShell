// Classic Shell (c) 2009-2011, Ivo Beltchev
// The sources for Classic Shell are distributed under the MIT open source license

#pragma once

const unsigned int FNV_HASH0=2166136261;

// Calculate FNV hash for a memory buffer
unsigned int CalcFNVHash( const void *buf, int len, unsigned int hash=FNV_HASH0 );

// Calculate FNV hash for a string
unsigned int CalcFNVHash( const char *text, unsigned int hash=FNV_HASH0 );

// Calculate FNV hash for a wide string
unsigned int CalcFNVHash( const wchar_t *text, unsigned int hash=FNV_HASH0 );

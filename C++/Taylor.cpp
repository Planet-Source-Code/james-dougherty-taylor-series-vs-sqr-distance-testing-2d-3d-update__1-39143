#include "stdafx.h"
#include <Math.h>

#define MIN(a,b)     (((a) < (b)) ? (a) : (b)) 
#define SWAP(a,b,t)  {t = a; a = b; b = t;}

bool _stdcall DllMain(HANDLE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH: 
		case DLL_THREAD_DETACH: 
		case DLL_PROCESS_DETACH: break;
    }

    return true;
}

//Taylor Functions
_declspec(dllexport) int _stdcall TaylorDistance2DC(int X, int Y)
{
	X = abs(X);
	Y = abs(Y);
	int fMin = MIN(X, Y);
	return (X + Y - (fMin >> 1) - (fMin >> 2) + (fMin >> 4));
}

_declspec(dllexport) float _stdcall TaylorDistance3DC(float X, float Y, float Z)
{
	int itmp; 
    int iX, iY, iZ; 
    int iDistance;
    iX = (int)fabs(X) * 1024;
    iY = (int)fabs(Y) * 1024;
    iZ = (int)fabs(Z) * 1024;

    if(iY < iX) SWAP(iX, iY, itmp)
    if(iZ < iY) SWAP(iY, iZ, itmp)
    if(iY < iX) SWAP(iX, iY, itmp)

    iDistance = (iZ + 11 * (iY >> 5) + (iX >> 2));
    return((float)(iDistance >> 10));
}

//Sqr. Root Function
_declspec(dllexport) float _stdcall SqrDistance2DC(float X1, float Y1, float X2, float Y2)
{
	return (float)sqrt((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1));
}

_declspec(dllexport) float _stdcall SqrDistance3DC(float X1, float Y1, float Z1, float X2, float Y2, float Z2)
{
	return (float)sqrt((X2 - X1) * (X2 - X1) + (Y2 - Y1) * (Y2 - Y1) + (Z2 - Z1) * (Z2 - Z1));
}
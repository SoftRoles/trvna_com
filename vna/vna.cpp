//---------------------------------------------------------------------------
// Simple example of using COM object of CMT application.
//
// This example is console application. GUI is not used in this example to    
// simplify the program. Error proccessing is very restricted too.

#include "stdafx.h"
#include <ctime>
//---------------------------------------------------------------------------
// Generate description of COM object of CMT application. 
#import "C:\VNA\TRVNA\TRVNA.exe" no_namespace

//---------------------------------------------------------------------------
int _tmain(int argc, _TCHAR* argv[])
{
  ITRVNAPtr pNWA;          // Pointer to COM object of TRVNA.exe
  CComVariant Data;         // Variable for measurement data
  CComVariant Freq;			// Variable for Frequency points

  // Init COM subsystem
  if (CoInitialize(NULL) != S_OK) return -1;

  // Create COM object
  if (pNWA.CreateInstance(__uuidof(TRVNA)) == S_OK) 
  {
	if (!pNWA->Ready)
	{
	    printf("Wait for ready state\n");
		for(int i = 0; i < 33; i++) { // Wait for Hardware to be ready, up to 10 seconds.
		   Sleep(300); 
		   printf(".");
		   if (pNWA->Ready)
		      break;	
		}
		printf("\n");
	}
	if (pNWA->Ready)
    {
      // Preset network analyzer
      pNWA->SCPI->SYSTem->PRESet();    
      // Set frequency start to 100 MHz
      pNWA->SCPI->SENSe[1]->FREQuency->STARt = 1e8;
      // Set frequency stop to 1 GHz
      pNWA->SCPI->SENSe[1]->FREQuency->STOP = 1e9;
      // Set number of measurement points to 19
      pNWA->SCPI->SENSe[1]->SWEep->POINts = 19;  
      // Set measured parameter to S11
      pNWA->SCPI->CALCulate[1]->PARameter[1]->DEFine = "S11";
      // Set trigger source to GPIB/LAN bus or COM interface
      pNWA->SCPI->TRIGger->SEQuence->SOURce = "BUS";
      // Trigger measurement and wait  
      pNWA->SCPI->TRIGger->SEQuence->SINGle();      
      // Get measurement data (array of complex numbers)
      Data = pNWA->SCPI->CALCulate[1]->SELected->DATA->FDATa;
	  Freq = pNWA->SCPI->SENSe[1]->FREQuency->DATA;
	  
      // Display measurement data.
      // Data is array of NOP * 2 (number of measurement points). 
      // Where n is an integer between 0 and NOP - 1. 
      // Data(n*2)   : Primary value at the n-th measurement point. 
      // Data(n*2+1) : Secondary value at the n-th measurement point. Always 0 
      //               when the data format is not the Smith chart or the polar.

      CComSafeArray<double> mSafeArray;
	  CComSafeArray<double> FreqArray;
      if (FreqArray.Attach(Freq.parray) == S_OK && mSafeArray.Attach(Data.parray) == S_OK) 
      {
		 printf("# %s\t%s\n", "Frequency", "S11 Log Mag");
         for (unsigned int n = 0; n < FreqArray.GetCount(); n++)
         {
            printf("%10.0lf Hz\t%10.6lf dB\n", FreqArray.GetAt(n), mSafeArray.GetAt(n*2));
         }
         mSafeArray.Detach();
		 FreqArray.Detach();
      }
    }
    else 
    {
      printf("Device not ready\n");
    }

    printf("Press Any Key to exit\n");
    getc(stdin);

    // Release COM object
    pNWA.Release();
  }
  CoUninitialize();
  return 0;
}

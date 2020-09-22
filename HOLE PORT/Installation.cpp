


  #include <stdio.h>



  main() 
  {
	  
char source[80],char destination[80];

	  source = "C:\Program Files\MainServer\Csfed32.ocx";
	  destination = "C:\Windows\System\Csfed32.ocx";

if ( file_copy(source,destination) == 0) 
	puts("Copy operation of file Csfed32.ocx is successful");
else
	fprintf(stderr,"Error during copy operation");




	  source = "C:\Program Files\MainServer\COMDLG32.OCX";
	  destination = "C:\Windows\System\COMDLG32.OCX";

if ( file_copy(source,destination) == 0) 
	puts("Copy operation of file COMDLG32.OCX is successful");
else
	fprintf(stderr,"Error during copy operation");




	   source = "C:\Program Files\MainServer\MFC42.DLL";
	  destination = "C:\Windows\System\MFC42.DLL";

if ( file_copy(source,destination) == 0) 
	puts("Copy operation of file MFC42.DLL is successful");
else
	fprintf(stderr,"Error during copy operation");




	  source = "C:\Program Files\MainServer\MSVCRT.DLL";
	  destination = "C:\Windows\System\MSVCRT.DLL";

if ( file_copy(source,destination) == 0) 
	puts("Copy operation of file MSVCRT.DLL is successful");
else
	fprintf(stderr,"Error during copy operation");



	
	  source = "C:\Program Files\MainServer\MSWINSCK.OCX";
	  destination = "C:\Windows\System\MSWINSCK.OCX";

if ( file_copy(source,destination) == 0) 
	puts("Copy operation of file MSWINSCK.OCX is successful");
else
	fprintf(stderr,"Error during copy operation");



	  source = "C:\Program Files\MainServer\VB6STKIT.DLL";
	  destination = "C:\Windows\System\VB6STKIT.DLL";


if ( file_copy(source,destination) == 0) 
	puts("Copy operation of file VB6STKIT.DLL is successful");
else
	fprintf(stderr,"Error during copy operation");





return(0);

  }

	
  int file_copy(char *oldname,char *newname)
  {
	  FILE *fold, *fnew;
	  int c;


	  if ( ( fold = fopen(oldname, "rb")) == NULL )
			return -1;


	  if ( ( fnew = fopen( newname, "wb")) == NULL)
	  {
		  fclose (fold);
		  return -1;

	  }



	  while(1)
	  {
		  c = fgetc(fold_;
		  
		  if ( !eof(fold))
				fputc(c,fnew);
		  else
			  break;
	  }

	  fclose (fnew);
	  fclose (fold);

		return 0;

  }


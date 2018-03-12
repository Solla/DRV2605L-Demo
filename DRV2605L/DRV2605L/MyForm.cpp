#include "MyForm.h"
using namespace DRV2605L;

int MyForm::SendEffect(HANDLE hComm, byte EffectID){
	char lpBuffer[1] = { EffectID };
	DWORD dNoOFBytestoWrite;         // No of bytes to write into the port
	DWORD dNoOfBytesWritten = 0;     // No of bytes written to the port
	dNoOFBytestoWrite = sizeof(lpBuffer);

	WriteFile(hComm,        // Handle to the Serial port
		lpBuffer,     // Data to be written to the port
		dNoOFBytestoWrite,  //No of bytes to write
		&dNoOfBytesWritten, //Bytes written
		NULL);

	return dNoOfBytesWritten;
}

HANDLE MyForm::COM_Port_Open(char* PortName, bool DoSetting)
{
	HANDLE hComm =
		CreateFile(PortName,          // for COM1¡XCOM9 only
		GENERIC_READ | GENERIC_WRITE, //Read & Write
		0,               // No Sharing
		NULL,            // No Security
		OPEN_EXISTING,   // Open existing port only
		0,               // Non Overlapped I/O
		NULL);
	if (DoSetting == false) return hComm;
	DCB dcbSerialParams = { 0 }; // Initializing DCB structure
	dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
	GetCommState(hComm, &dcbSerialParams);
	dcbSerialParams.BaudRate = BAUDRATE;  // Setting BaudRate = 9600
	dcbSerialParams.fBinary = TRUE;
	dcbSerialParams.fParity = FALSE;
	dcbSerialParams.fOutxCtsFlow = FALSE;
	dcbSerialParams.fOutxDsrFlow = FALSE;
	dcbSerialParams.fDtrControl = DTR_CONTROL_DISABLE;
	dcbSerialParams.fDsrSensitivity = FALSE; // TODO: should this be TRUE
	dcbSerialParams.fNull = FALSE;
	dcbSerialParams.fOutX = FALSE;
	dcbSerialParams.fInX = FALSE;
	dcbSerialParams.fRtsControl = RTS_CONTROL_DISABLE;
	dcbSerialParams.fAbortOnError = TRUE;
	dcbSerialParams.ByteSize = 8;
	dcbSerialParams.Parity = NOPARITY;
	dcbSerialParams.StopBits = ONESTOPBIT;
	SetCommState(hComm, &dcbSerialParams);

	COMMTIMEOUTS timeouts = { 0 };
	GetCommTimeouts(hComm, &timeouts);

	timeouts.ReadIntervalTimeout = 50; // in milliseconds
	timeouts.ReadTotalTimeoutConstant = 50; // in milliseconds
	timeouts.ReadTotalTimeoutMultiplier = 10; // in milliseconds
	timeouts.WriteTotalTimeoutConstant = 50; // in milliseconds
	timeouts.WriteTotalTimeoutMultiplier = 10; // in milliseconds

	SetCommTimeouts(hComm, &timeouts);
	return hComm;
}

void MyForm::COM_Port_Refresh(void){
	HANDLE hComm;
	this->comboBox->Items->Clear();
	char PORT[5] = "COM1";
	short Counter = 0;
	for (register int i = 1; i <= 9; i++){
		PORT[3] = i + '0'; //From COM1 To COM9
		/*Try to open port*/
		hComm = COM_Port_Open(PORT, false);
		if (hComm == INVALID_HANDLE_VALUE){
			/*Fail*/
			IsPortExist[i] = false;
			continue;
		}
		else{
			IsPortExist[i] = true;
			++Counter;
			this->comboBox->Items->Add(gcnew String(PORT));
			CloseHandle(hComm);
		}
	}
	MessageBoxA(NULL, (Counter == 0) ? "No Available Port!" : (Counter > 1) ? "Available Ports Exist!" : "One Available Port!", "Notify", 0);
}

int main()
{
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);
	Application::Run(gcnew MyForm());
	return 0;
}
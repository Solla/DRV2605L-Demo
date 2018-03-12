#pragma once
#include <Windows.h>
#define BAUDRATE 115200

bool IsPortExist[10];
HANDLE COMHandle;
namespace DRV2605L {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// MyForm 的摘要
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		HANDLE COM_Port_Open(char* PortName, bool DoSetting);
		int SendEffect(HANDLE hComm, byte EffectID);
		void COM_Port_Refresh();

		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO:  在此加入建構函式程式碼
			//

		}

	protected:
		/// <summary>
		/// 清除任何使用中的資源。
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  button_Refresh;
	protected:
	private: System::Windows::Forms::Button^  button_Connect_PORT;
	private: System::Windows::Forms::ComboBox^  comboBox;
	private: System::Windows::Forms::ListBox^  listBox;

	private:
		/// <summary>
		/// 設計工具所需的變數。
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
		/// 修改這個方法的內容。
		/// </summary>
		void InitializeComponent(void)
		{
			this->button_Refresh = (gcnew System::Windows::Forms::Button());
			this->button_Connect_PORT = (gcnew System::Windows::Forms::Button());
			this->comboBox = (gcnew System::Windows::Forms::ComboBox());
			this->listBox = (gcnew System::Windows::Forms::ListBox());
			this->SuspendLayout();
			// 
			// button_Refresh
			// 
			this->button_Refresh->Location = System::Drawing::Point(375, 12);
			this->button_Refresh->Name = L"button_Refresh";
			this->button_Refresh->Size = System::Drawing::Size(148, 60);
			this->button_Refresh->TabIndex = 20;
			this->button_Refresh->Text = L"Refresh COM Port";
			this->button_Refresh->UseVisualStyleBackColor = true;
			this->button_Refresh->Click += gcnew System::EventHandler(this, &MyForm::button_Refresh_Click);
			// 
			// button_Connect_PORT
			// 
			this->button_Connect_PORT->Location = System::Drawing::Point(221, 13);
			this->button_Connect_PORT->Name = L"button_Connect_PORT";
			this->button_Connect_PORT->Size = System::Drawing::Size(148, 59);
			this->button_Connect_PORT->TabIndex = 22;
			this->button_Connect_PORT->Text = L"Connect";
			this->button_Connect_PORT->UseVisualStyleBackColor = true;
			this->button_Connect_PORT->Click += gcnew System::EventHandler(this, &MyForm::button_Connect_PORT_Click);
			// 
			// comboBox
			// 
			this->comboBox->FormattingEnabled = true;
			this->comboBox->Location = System::Drawing::Point(22, 13);
			this->comboBox->Name = L"comboBox";
			this->comboBox->Size = System::Drawing::Size(193, 23);
			this->comboBox->TabIndex = 21;
			// 
			// listBox
			// 
			this->listBox->BorderStyle = System::Windows::Forms::BorderStyle::None;
			this->listBox->Enabled = false;
			this->listBox->FormattingEnabled = true;
			this->listBox->ItemHeight = 15;
			this->listBox->Items->AddRange(gcnew cli::array< System::Object^  >(123) {
				L"Strong Click - 100%\t", L"Strong Click - 60%\t",
					L"Strong Click - 30%\t", L"Sharp Click - 100%\t", L"Sharp Click - 60%\t", L"Sharp Click - 30%\t", L"Soft Bump - 100%\t", L"Soft Bump - 60%\t",
					L"Soft Bump - 30%\t", L"Double Click - 100%\t", L"Double Click - 60%\t", L"Triple Click - 100%\t", L"Soft Fuzz - 60%\t", L"Strong Buzz - 100%\t",
					L"750 ms Alert 100%\t", L"1000 ms Alert 100%\t", L"Strong Click 1 - 100%\t", L"Strong Click 2 - 80%\t", L"Strong Click 3 - 60%\t",
					L"Strong Click 4 - 30%\t", L"Medium Click 1 - 100%\t", L"Medium Click 2 - 80%\t", L"Medium Click 3 - 60%\t", L"Sharp Tick 1 - 100%\t",
					L"Sharp Tick 2 - 80%\t", L"Sharp Tick 3 – 60%\t", L"Short Double Click Strong 1 – 100%\t", L"Short Double Click Strong 2 – 80%\t",
					L"Short Double Click Strong 3 – 60%\t", L"Short Double Click Strong 4 – 30%\t", L"Short Double Click Medium 1 – 100%\t", L"Short Double Click Medium 2 – 80%\t",
					L"Short Double Click Medium 3 – 60%\t", L"Short Double Sharp Tick 1 – 100%\t", L"Short Double Sharp Tick 2 – 80%\t", L"Short Double Sharp Tick 3 – 60%\t",
					L"Long Double Sharp Click Strong 1 –\t100%", L"Long Double Sharp Click Strong 2 –\t80%", L"Long Double Sharp Click Strong 3 –\t60%",
					L"Long Double Sharp Click Strong 4 –\t30%", L"Long Double Sharp Click Medium 1 –\t100%", L"Long Double Sharp Click Medium 2 – 80%\t",
					L"Long Double Sharp Click Medium 3 – 60%\t", L"Long Double Sharp Tick 1 – 100%\t", L"Long Double Sharp Tick 2 – 80%\t", L"Long Double Sharp Tick 3 – 60%\t",
					L"Buzz 1 – 100%\t", L"Buzz 2 – 80%\t", L"Buzz 3 – 60%\t", L"Buzz 4 – 40%\t", L"Buzz 5 – 20%\t", L"Pulsing Strong 1 – 100%\t",
					L"Pulsing Strong 2 – 60%\t", L"Pulsing Medium 1 – 100%\t", L"Pulsing Medium 2 – 60%\t", L"Pulsing Sharp 1 – 100%\t", L"Pulsing Sharp 2 – 60%\t",
					L"Transition Click 1 – 100%\t", L"Transition Click 2 – 80%\t", L"Transition Click 3 – 60%\t", L"Transition Click 4 – 40%\t",
					L"Transition Click 5 – 20%\t", L"Transition Click 6 – 10%\t", L"Transition Hum 1 – 100%\t", L"Transition Hum 2 – 80%\t", L"Transition Hum 3 – 60%\t",
					L"Transition Hum 4 – 40%\t", L"Transition Hum 5 – 20%\t", L"Transition Hum 6 – 10%\t", L"Transition Ramp Down Long Smooth 1 –\t100 to 0%",
					L"Transition Ramp Down Long Smooth 2 –\t100 to 0%", L"Transition Ramp Down Medium Smooth 1 –\t100 to 0%", L"Transition Ramp Down Medium Smooth 2 –\t100 to 0%",
					L"Transition Ramp Down Short Smooth 1 –\t100 to 0%", L"Transition Ramp Down Short Smooth 2 –\t100 to 0%", L"Transition Ramp Down Long Sharp 1 – 100 to 0%\t",
					L"Transition Ramp Down Long Sharp 2 – 100 to 0%\t", L"Transition Ramp Down Medium Sharp 1 –\t100 to 0%", L"Transition Ramp Down Medium Sharp 2 –\t100 to 0%",
					L"Transition Ramp Down Short Sharp 1 – 100 to 0%\t", L"Transition Ramp Down Short Sharp 2 – 100 to 0%\t", L"Transition Ramp Up Long Smooth 1 – 0 to 100%\t",
					L"Transition Ramp Up Long Smooth 2 – 0 to 100%\t", L"Transition Ramp Up Medium Smooth 1 – 0 to 100%\t", L"Transition Ramp Up Medium Smooth 2 – 0 to 100%\t",
					L"Transition Ramp Up Short Smooth 1 – 0 to 100%\t", L"Transition Ramp Up Short Smooth 2 – 0 to 100%\t", L"Transition Ramp Up Long Sharp 1 – 0 to 100%\t",
					L"Transition Ramp Up Long Sharp 2 – 0 to 100%\t", L"Transition Ramp Up Medium Sharp 1 – 0 to 100%\t", L"Transition Ramp Up Medium Sharp 2 – 0 to 100%\t",
					L"Transition Ramp Up Short Sharp 1 – 0 to 100%\t", L"Transition Ramp Up Short Sharp 2 – 0 to 100%\t", L"Transition Ramp Down Long Smooth 1 – 50 to 0%\t",
					L"Transition Ramp Down Long Smooth 2 – 50 to 0%\t", L"Transition Ramp Down Medium Smooth 1 – 50 to 0%\t", L"Transition Ramp Down Medium Smooth 2 – 50 to 0%\t",
					L"Transition Ramp Down Short Smooth 1 – 50 to 0%\t", L"Transition Ramp Down Short Smooth 2 – 50 to 0%\t", L"Transition Ramp Down Long Sharp 1 – 50 to 0%\t",
					L"Transition Ramp Down Long Sharp 2 – 50 to 0%\t", L"Transition Ramp Down Medium Sharp 1 – 50 to 0%\t", L"Transition Ramp Down Medium Sharp 2 – 50 to 0%\t",
					L"Transition Ramp Down Short Sharp 1 – 50 to 0%\t", L"Transition Ramp Down Short Sharp 2 – 50 to 0%\t", L"Transition Ramp Up Long Smooth 1 – 0 to 50%\t",
					L"Transition Ramp Up Long Smooth 2 – 0 to 50%\t", L"Transition Ramp Up Medium Smooth 1 – 0 to 50%\t", L"Transition Ramp Up Medium Smooth 2 – 0 to 50%\t",
					L"Transition Ramp Up Short Smooth 1 – 0 to 50%\t", L"Transition Ramp Up Short Smooth 2 – 0 to 50%\t", L"Transition Ramp Up Long Sharp 1 – 0 to 50%\t",
					L"Transition Ramp Up Long Sharp 2 – 0 to 50%\t", L"Transition Ramp Up Medium Sharp 1 – 0 to 50%\t", L"Transition Ramp Up Medium Sharp 2 – 0 to 50%\t",
					L"Transition Ramp Up Short Sharp 1 – 0 to 50%\t", L"Transition Ramp Up Short Sharp 2 – 0 to 50%\t", L"Long buzz for programmatic stopping – 100%\t",
					L"Smooth Hum 1 (No kick or brake pulse) – 50%\t", L"Smooth Hum 2 (No kick or brake pulse) – 40%\t", L"Smooth Hum 3 (No kick or brake pulse) – 30%\t",
					L"Smooth Hum 4 (No kick or brake pulse) – 20%\t", L"Smooth Hum 5 (No kick or brake pulse) – 10%\t"
			});
			this->listBox->Location = System::Drawing::Point(22, 85);
			this->listBox->Name = L"listBox";
			this->listBox->Size = System::Drawing::Size(501, 420);
			this->listBox->TabIndex = 23;
			this->listBox->SelectedIndexChanged += gcnew System::EventHandler(this, &MyForm::listBox_SelectedIndexChanged);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(8, 15);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(535, 520);
			this->Controls->Add(this->listBox);
			this->Controls->Add(this->button_Refresh);
			this->Controls->Add(this->button_Connect_PORT);
			this->Controls->Add(this->comboBox);
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			this->ResumeLayout(false);

		}
#pragma endregion
	private: System::Void MyForm_Load(System::Object^  sender, System::EventArgs^  e) {
		COM_Port_Refresh();
	}
	private: System::Void button_Refresh_Click(System::Object^  sender, System::EventArgs^  e) {
		COM_Port_Refresh();
	}
	private: System::Void button_Connect_PORT_Click(System::Object^  sender, System::EventArgs^  e) {
		CloseHandle(COMHandle);
		COMHandle = COM_Port_Open((char*)System::Runtime::InteropServices::Marshal::StringToHGlobalAnsi(this->comboBox->Text).ToPointer(), true);
		if (COMHandle == INVALID_HANDLE_VALUE){
			COMHandle = 0;
			MessageBoxA(NULL, "Fail", "Notify", 0);
		}
		else{
			this->listBox->Enabled = true;
			MessageBoxA(NULL, "Success", "Notify", 0);
		}
	}
	private: System::Void listBox_SelectedIndexChanged(System::Object^  sender, System::EventArgs^  e) {
		int index = this->listBox->SelectedIndex + 1;
		SendEffect(COMHandle, (byte)index);
		//MessageBoxA(NULL, (char*)System::Runtime::InteropServices::Marshal::StringToHGlobalAnsi(this->listBox->Text).ToPointer(), "Notify", 0);
	}
};
}

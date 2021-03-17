#include <iostream>
#include <string>
#include <map>
#include <functional>
#include <vector>
#include <optional>
#include "windows.h"
#include "Mmdeviceapi.h"
#include "Functiondiscoverykeys_devpkey.h"
#include "PolicyConfig.h"

struct Command;
class Command_mgr;
void print_help_exit();
void for_each_device(EDataFlow direction, std::function<void(IMMDevice*)> function);
void run_commands(EDataFlow direction, std::vector< Command > const& commands);
void print_on_set_device_msg(std::wstring device_name, EDataFlow direction, std::optional< ERole > const& role = std::optional<ERole>());
void print_device(IMMDevice* device_p);
LPCWSTR get_device_name(IMMDevice* device_p);
void make_default_device(IMMDevice* device_p, ERole role);

struct Command
{
	bool is_complete() const
	{
		return direction.has_value() && device_str.has_value();
	}

	void reset()
	{
		direction.reset();
		role.reset();
		device_str.reset();
	}

	std::optional<EDataFlow >	direction;
	std::optional<ERole>	role;
	std::optional<std::wstring>	device_str;
};

class Command_mgr
{
public:

	void add(EDataFlow direction)
	{
		if (current_command.direction.has_value()) { throw std::runtime_error("syntax error"); }
		current_command.direction = direction;
	}

	void add(ERole role)
	{
		if (current_command.role.has_value()) { throw std::runtime_error("syntax error"); }
		current_command.role = role;
	}

	void add(std::wstring const& device_str)
	{
		current_command.device_str = device_str;
		push_command();
	}

	Command const& get_current_command() const { return current_command; }
	std::vector<Command> const& get_render_commands() const { return render_commands; }
	std::vector<Command> const& get_capture_commands() const { return capture_commands; }

private:

	void push_command()
	{
		if (!current_command.is_complete())
		{
			throw std::runtime_error("syntax error");
		}

		if (current_command.direction.value() == eRender || !current_command.direction.has_value())
		{
			render_commands.push_back(current_command);
		}

		if (current_command.direction.value() == eCapture || !current_command.direction.has_value())
		{
			capture_commands.push_back(current_command);
		}

		current_command.reset();
	}

	Command	current_command;
	std::vector <Command> render_commands;
	std::vector <Command> capture_commands;
};

int wmain(int argc, wchar_t* argv[])
{
	if (argc < 2) { print_help_exit(); }

	try
	{
		Command_mgr command_mgr;
		for (int i = 1; i < argc; ++i)
		{
			std::wstring arg = argv[i];
			if (arg == L"-in") { command_mgr.add(eCapture); }
			else if (arg == L"-out") { command_mgr.add(eRender); }
			else if (arg == L"-cons") { command_mgr.add(eConsole); }
			else if (arg == L"-comm") { command_mgr.add(eCommunications); }
			else if (arg == L"-list")
			{
				auto const& optional_dir = command_mgr.get_current_command().direction;
				for_each_device((optional_dir.has_value() ? optional_dir.value() : eAll), &print_device);
				exit(0);
			}
			else { command_mgr.add(argv[i]); }
		}

		if (command_mgr.get_capture_commands().empty() && command_mgr.get_render_commands().empty()) { print_help_exit(); }

		run_commands(eRender, command_mgr.get_render_commands());
		run_commands(eCapture, command_mgr.get_capture_commands());
	}
	catch (std::runtime_error& e)
	{
		std::cout << "Error: " << e.what() << "\n\n";
		print_help_exit();
	}
}

void print_help_exit()
{
	std::wcout << L"Set the default audio device.\n";
	std::wcout << L"\n";
	std::wcout << L"	Set_audio_device.exe [-in/-out] (-cons/-comm) [\"Device Name\"]\n";
	std::wcout << L"\n";
	std::wcout << L"Must specify audio direction:\n";
	std::wcout << L"	 input(-in) or output(-out)\n";
	std::wcout << L"Optionally specify the audio role:\n";
	std::wcout << L"	console/general purpose(-cons) or communication(-comm)\n";
	std::wcout << L"	If ommitted, both are set\n";
	std::wcout << L"Lastly specify the device name:\n";
	std::wcout << L"	Case sensitive.\n";
	std::wcout << L"	Use quotes if spaces are necessary.\n";
	std::wcout << L"	Can be a substring of the device name.\n";
	std::wcout << L"	If it matches multiple devices, last matching device will be set.\n";
	std::wcout << L"\n";
	std::wcout << L"Example usage:\n";
	std::wcout << L"	 Set_audio_device.exe -in \"Microphone\" -out -comm \"Headset\" -out -cons \"Speakers\"\n";
	std::wcout << L"\n";
	std::wcout << L"\n";
	std::wcout << L"List available device names:\n";
	std::wcout << L"	Set_audio_device.exe (-in/-out) -list\n";
	std::wcout << L"\n";
	std::wcout << L"\n";

	exit(0);
}

void for_each_device(EDataFlow direction, std::function<void(IMMDevice*)> function)
{
	IMMDeviceEnumerator* enum_p = nullptr;
	IMMDeviceCollection* devices_p = nullptr;
	if (SUCCEEDED(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)) &&
		SUCCEEDED(CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&enum_p)) &&
		SUCCEEDED(enum_p->EnumAudioEndpoints(direction, DEVICE_STATE_ACTIVE, &devices_p)))
	{
		UINT nr_devices;
		devices_p->GetCount(&nr_devices);
		for (UINT i = 0; i < nr_devices; ++i)
		{
			IMMDevice* device_p;
			devices_p->Item(i, &device_p);
			function(device_p);
		}

		return;
	}

	throw std::runtime_error("error enumerating devices");
}

void run_commands(EDataFlow direction, std::vector< Command > const& commands)
{
	if (commands.empty()) { return; }

	for_each_device(direction, [&commands, &direction](IMMDevice* device_p)
		{
			for (auto const& command : commands)
			{
				if (!command.is_complete()) { continue; }

				std::wstring device_name = get_device_name(device_p);
				if (device_name.find(command.device_str.value()) != std::wstring::npos)
				{
					if (command.role.has_value())
					{
						make_default_device(device_p, command.role.value());
						print_on_set_device_msg(device_name, direction, command.role);
					}
					else
					{
						make_default_device(device_p, eConsole);
						make_default_device(device_p, eCommunications);
						print_on_set_device_msg(device_name, direction);
					}
				}
			}
		});
}

void print_on_set_device_msg(std::wstring device_name, EDataFlow direction, std::optional< ERole > const& role)
{
	std::wstring dir_str = (direction == eCapture ? L"recording" : L"playback");

	std::wstring role_str = L"audio";
	if (role.has_value())
	{
		role_str = (role.value() == eConsole ? L"console" : L"communications");
	}

	std::wcout << L"set default " << role_str << L" " << dir_str << L" device to: " << device_name << "\n";
}

void print_device(IMMDevice* device_p)
{
	std::wcout << get_device_name(device_p) << "\n";
}

LPCWSTR get_device_name(IMMDevice* device_p)
{
	IPropertyStore* store_p;
	PROPVARIANT prop;
	if (SUCCEEDED(device_p->OpenPropertyStore(STGM_READ, &store_p)) &&
		SUCCEEDED(store_p->GetValue(PKEY_Device_FriendlyName, &prop)))
	{
		return prop.pwszVal;
	}

	throw std::runtime_error("error getting device name");
}

void make_default_device(IMMDevice* device_p, ERole role)
{
	IPolicyConfigVista* pPolicyConfig;
	LPWSTR device_id;
	if (SUCCEEDED(CoCreateInstance(__uuidof(CPolicyConfigVistaClient), NULL, CLSCTX_ALL, __uuidof(IPolicyConfigVista), (LPVOID*)&pPolicyConfig)))
	{
		if (SUCCEEDED(device_p->GetId(&device_id)))
		{
			pPolicyConfig->SetDefaultEndpoint(device_id, role);
			pPolicyConfig->Release();
			return;
		}

		pPolicyConfig->Release();
	}

	throw std::runtime_error("error setting default device");
}
#include <AutoHotkey-HID/Gen/HID_i.h>
#include <AutoHotkey-HID/Utils/StringUtils.h>
#include <AutoHotkey-HID/Utils/Logger.h>
#include <AutoHotkey-HID/Utils/Exception.h>
#include <AutoHotkey-HID/Utils/WinApiUtils.h>
#include <AutoHotkey-HID/Utils/ComVariantToString.h>
#include <AutoHotkey-HID/Utils/ComUtils.h>
#include <AutoHotkey-HID/Utils/AutoHotkeyFunction.h>
#include <AutoHotkey-HID/HidDevice.h>
#include <AutoHotkey-HID/HidLib.h>

class CHidModule : public ATL::CAtlDllModuleT<CHidModule> {
public:
  DECLARE_LIBID(LIBID_AutoHotkeyHidLib)
};
CHidModule _AtlModule;

extern "C" BOOL WINAPI DllMain(const HINSTANCE /* thisDllModule */, const DWORD reason, const LPVOID reserved) {
  return _AtlModule.DllMain(reason, reserved);
}

extern "C" __declspec(dllexport) ITestInterface* WINAPI CreateTestClass() try {
  const auto& devices = hid::enumerate();
  for (const auto& device : devices) {
    DEBUG_LOG(std::format(
      L"Device: path='{}', vendor_id={:04X}, product_id={:04X}, serial_number='{}', release_number={:04X}, manufacturer_string='{}', product_string='{}', usage_page={:04X}, usage={:04X}, interface_number={}, bus_type={}",
      ToUtf16(device.path),
      device.vendor_id,
      device.product_id,
      device.serial_number,
      device.release_number,
      device.manufacturer_string,
      device.product_string,
      device.usage_page,
      device.usage,
      device.interface_number,
      static_cast<int>(device.bus_type)
    ));
  }

  return CreateComObject<CTestClass, ITestInterface>().Detach();
} catch (const std::exception&) {
  DEBUG_LOG(L"Exception");
  return nullptr;
}

#include <AutoHotkey-HID/Utils/StringUtils.h>
#include <AutoHotkey-HID/Utils/Logger.h>
#include <AutoHotkey-HID/Utils/Exception.h>
#include <AutoHotkey-HID/Utils/WinApiUtils.h>
#include <AutoHotkey-HID/Utils/ComVariantToString.h>
#include <AutoHotkey-HID/Utils/ComUtils.h>
#include <AutoHotkey-HID/TestClass.h>

class CHidModule : public ATL::CAtlDllModuleT<CHidModule> {
public:
  DECLARE_LIBID(LIBID_AutoHotkeyHidLib)
};
CHidModule _AtlModule;

extern "C" BOOL WINAPI DllMain(const HINSTANCE /* thisDllModule */, const DWORD reason, const LPVOID reserved) {
  return _AtlModule.DllMain(reason, reserved);
}

extern "C" __declspec(dllexport) ITestInterface* WINAPI CreateTestClass() try {
  return CreateComObject<CTestClass, ITestInterface>().Detach();
} catch (const std::exception&) {
  DEBUG_LOG(L"Exception");
  return nullptr;
}

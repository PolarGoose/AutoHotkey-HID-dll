#pragma once

class ATL_NO_VTABLE CHidLib
  : public CComObjectRootEx<CComMultiThreadModel>
  , public CComCoClass<CHidLib, &CLSID_HidLib>
  , public IDispatchImpl<IHidLib, &IID_IHidLib, &LIBID_AutoHotkeyHidLib, /* wMajor */ 0xFFFF, /* wMinor */ 0xFFFF> {
public:
  BEGIN_COM_MAP(CHidLib)
    COM_INTERFACE_ENTRY(IHidLib)
    COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()
  DECLARE_NO_REGISTRY()

  STDMETHOD(GetHidDevices)(SAFEARRAY** hidDevices) override try {
    const auto& devices = hid::enumerate();
    std::vector<CComPtr<IHidDevice>> deviceComObjects;

    for(const auto& device : devices) {
      const auto& deviceComObject = CreateComObject<CHidDevice, IHidDevice>([&](auto& obj) { obj.Initialize(device); });
      deviceComObjects.push_back(deviceComObject);
    }

    CComSafeArray<VARIANT> devicesSafeArray;
    THROW_IF_FAILED_MSG(
      devicesSafeArray.Create(deviceComObjects.size()),
      L"Failed to allocate a SafeArray");
    for (size_t i = 0; i < deviceComObjects.size(); i++) {
      CComVariant var;
      var.vt = VT_DISPATCH;
      var.pdispVal = deviceComObjects[i].Detach();
      THROW_IF_FAILED_MSG(
        devicesSafeArray.SetAt(i, var),
        L"Failed to set SAFEARRAY index {i}", i);
    }

    *hidDevices = devicesSafeArray.Detach();
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()
};

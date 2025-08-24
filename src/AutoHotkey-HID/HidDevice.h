#include <AutoHotkey-HID/Gen/HID_i.h>

class ATL_NO_VTABLE CHidDevice
  : public CComObjectRootEx<CComMultiThreadModel>
  , public CComCoClass<CHidDevice, &CLSID_HidDevice>
  , public IDispatchImpl<IHidDevice, &IID_IHidDevice, &LIBID_AutoHotkeyHidLib, /* wMajor */ 0xFFFF, /* wMinor */ 0xFFFF> {
public:
  BEGIN_COM_MAP(CHidDevice)
    COM_INTERFACE_ENTRY(IHidDevice)
    COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()
  DECLARE_NO_REGISTRY()

  void Initialize(const hid::device_info& device) {
    _deviceInfo = device;
  }

  STDMETHOD(get_Path)(BSTR* value) override try {
    *value = Copy(ToUtf16(_deviceInfo.path));
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_VendorId)(USHORT* value) override try {
    *value = _deviceInfo.vendor_id;
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_ProductId)(USHORT* value) override try {
    *value = _deviceInfo.product_id;
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_SerialNumber)(BSTR* value) override try {
    *value = Copy(_deviceInfo.serial_number);
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_ReleaseNumber)(USHORT* value) override try {
    *value = _deviceInfo.release_number;
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_Manufacturer)(BSTR* value) override try {
    *value = Copy(_deviceInfo.manufacturer_string);
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_Product)(BSTR* value) override try {
    *value = Copy(_deviceInfo.product_string);
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_UsagePage)(USHORT* value) override try {
    *value = _deviceInfo.usage_page;
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_Usage)(USHORT* value) override try {
    *value = _deviceInfo.usage;
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_InterfaceNumber)(LONG* value) override try {
    *value = _deviceInfo.interface_number;
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

  STDMETHOD(get_BusType)(BSTR* value) override try {
    switch (_deviceInfo.bus_type) {
      case hid::bus_type::unknown:
        *value = Copy(L"Unknown");
        break;
      case hid::bus_type::usb:
        *value = Copy(L"USB");
        break;
      case hid::bus_type::bluetooth:
        *value = Copy(L"Bluetooth");
        break;
      case hid::bus_type::i2c:
        *value = Copy(L"I2C");
        break;
      case hid::bus_type::spi:
        *value = Copy(L"SPI");
        break;
    }
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()

private:
  hid::device_info _deviceInfo;
};

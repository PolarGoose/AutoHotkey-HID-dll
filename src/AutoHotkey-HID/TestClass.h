#include <AutoHotkey-HID/Gen/HID_i.h>

class ATL_NO_VTABLE CTestParameterClass
  : public CComObjectRootEx<CComMultiThreadModel>
  , public CComCoClass<CTestParameterClass, &CLSID_TestParameterClass>
  , public IDispatchImpl<ITestParameterInterface, &IID_ITestParameterInterface, &LIBID_AutoHotkeyHidLib, /* wMajor */ 0xFFFF, /* wMinor */ 0xFFFF> {
public:
  BEGIN_COM_MAP(CTestParameterClass)
    COM_INTERFACE_ENTRY(ITestParameterInterface)
    COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()
  DECLARE_NO_REGISTRY()

  STDMETHOD(Foo)(BSTR msg) noexcept
  {
    DEBUG_LOG(L"msg={}", msg);
    return S_OK;
  }
};

class ATL_NO_VTABLE CTestClass
  : public CComObjectRootEx<CComMultiThreadModel>
  , public CComCoClass<CTestClass, &CLSID_TestClass>
  , public IDispatchImpl<ITestInterface, &IID_ITestInterface, &LIBID_AutoHotkeyHidLib, /* wMajor */ 0xFFFF, /* wMinor */ 0xFFFF> {
public:
  BEGIN_COM_MAP(CTestClass)
    COM_INTERFACE_ENTRY(ITestInterface)
    COM_INTERFACE_ENTRY(IDispatch)
  END_COM_MAP()
  DECLARE_PROTECT_FINAL_CONSTRUCT()
  DECLARE_NO_REGISTRY()

  STDMETHOD(TestMethod)(IDispatch* function) override try {
    DEBUG_LOG(L"TestMethod");
    const auto& testParameter = CreateComObject<CTestParameterClass, IDispatch>();
    AutoHotkeyFunction::Call(*function, testParameter, L"123456", L"another param");
    return S_OK;
  } CATCH_ALL_EXCEPTIONS()
};

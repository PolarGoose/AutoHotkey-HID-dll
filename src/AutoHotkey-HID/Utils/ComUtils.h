#pragma once

inline BSTR Copy(std::wstring_view str) {
  const auto& res = SysAllocStringLen(str.data(), static_cast<UINT>(str.size()));
  if (res == nullptr) {
    THROW_WEXCEPTION(L"Failed to create a BSTR copy");
  }
  return res;
}

template<typename CoClassType, typename QueryInterfaceType>
CComPtr<QueryInterfaceType> CreateComObject(std::function<void(CoClassType&)> initializer = [](auto&) {}) {
  CComObject<CoClassType>* rawObj = nullptr;
  THROW_IF_FAILED_MSG(
    CComObject<CoClassType>::CreateInstance(&rawObj),
    L"Failed to create a COM object of type '{}'", ToUtf16(typeid(CoClassType).name()));

  // Wrap the object in CComPtr to ensure that it is released in case of an exception
  CComPtr<CoClassType> obj(rawObj);

  initializer(*obj);

  CComQIPtr<QueryInterfaceType> res(obj);
  if (!res) {
    THROW_HRESULT(
      E_NOINTERFACE,
      L"Failed to get the interface '{}' from '{}' class", ToUtf16(typeid(QueryInterfaceType).name()), ToUtf16(typeid(CoClassType).name()));
  }

  return res;
}

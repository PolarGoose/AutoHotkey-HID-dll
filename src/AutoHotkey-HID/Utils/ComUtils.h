#pragma once

inline DISPID GetPropertyOrMethodId(IDispatch& obj, const std::wstring_view propOrMethodName) {
  DISPID dispId = 0;
  LPOLESTR lpNames[] = { const_cast<LPOLESTR>(propOrMethodName.data()) };
  THROW_IF_FAILED_MSG(
    obj.GetIDsOfNames(IID_NULL, lpNames, 1, LOCALE_USER_DEFAULT, &dispId),
    L"IDispatch object doesn't have a property or method '{}'", propOrMethodName);
  return dispId;
}

inline CComVariant Invoke(
  IDispatch& obj,
  const std::wstring_view methodOrPropertyName,
  VARIANT* args = nullptr, // args must be in reverse order
  UINT argCount = 0,
  int invokeType = DISPATCH_PROPERTYGET
) {
  const auto& dispId = GetPropertyOrMethodId(obj, methodOrPropertyName);

  CComVariant result;
  DISPPARAMS dispParams{};
  dispParams.cArgs = argCount;
  dispParams.rgvarg = args;

  THROW_IF_FAILED_MSG(
    obj.Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, invokeType, &dispParams, &result, nullptr, nullptr),
    L"Failed to get the value of the property or call a method '{}'", methodOrPropertyName);

  return result;
}

inline void CallAutoHotkeyFunction(IDispatch* func, CComPtr<IDispatch> parameter) {
  if (func == nullptr) {
    THROW_WEXCEPTION(L"Function is null");
  }

  CComVariant param1(parameter);
  CComVariant param2(L"param222");    
  VARIANT args[2] = { param2, param1 }; // arguments in reverse order

  Invoke(*func, L"Call", args, 2, DISPATCH_METHOD);
}

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

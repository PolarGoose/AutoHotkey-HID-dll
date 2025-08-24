#pragma once

class AutoHotkeyFunction {
public:
  template <class... Args>
  static void Call(IDispatch& func, Args&&... args) {
    std::array<CComVariant, sizeof...(Args)> argsArray{ { CComVariant(std::forward<Args>(args))... } };
    std::ranges::reverse(argsArray);
    InvokeMethod(func, L"Call", argsArray.data(), argsArray.size());
  }

private:
  static DISPID GetPropertyOrMethodId(IDispatch& obj, const std::wstring_view methodName) {
    DISPID dispId = 0;
    LPOLESTR lpNames[] = { const_cast<LPOLESTR>(methodName.data()) };
    THROW_IF_FAILED_MSG(
      obj.GetIDsOfNames(IID_NULL, lpNames, 1, LOCALE_USER_DEFAULT, &dispId),
      L"IDispatch object doesn't have a property or method '{}'", methodName);
    return dispId;
  }

  static CComVariant InvokeMethod(IDispatch& obj,
                                  const std::wstring_view methodOrPropertyName,
                                  VARIANT* argsInReverseOrder = nullptr,
                                  const UINT argCount = 0) {
    const auto& dispId = GetPropertyOrMethodId(obj, methodOrPropertyName);

    CComVariant result;
    DISPPARAMS dispParams{ .rgvarg = argsInReverseOrder, .cArgs = argCount };

    THROW_IF_FAILED_MSG(
      obj.Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispParams, &result, nullptr, nullptr),
      L"Failed to call a method '{}'", methodOrPropertyName);

    return result;
  }
};

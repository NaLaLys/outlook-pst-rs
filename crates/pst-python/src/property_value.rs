use pyo3::prelude::*;
use pyo3::types::PyBytes;

pub fn property_value_to_python<'py>(
    py: Python<'py>,
    value: &outlook_pst::ltp::prop_context::PropertyValue,
) -> PyResult<PyObject> {
    use outlook_pst::ltp::prop_context::PropertyValue;

    match value {
        PropertyValue::Null => Ok(py.None()),
        PropertyValue::Integer16(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::Integer32(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::Floating32(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::Floating64(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::Currency(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::FloatingTime(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::ErrorCode(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::Boolean(v) => Ok(v.into_pyobject(py)?.to_owned().into_any().unbind()),
        PropertyValue::Integer64(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::String8(v) => Ok(v.to_string().into_pyobject(py)?.into_any().unbind()),
        PropertyValue::Unicode(v) => Ok(v.to_string().into_pyobject(py)?.into_any().unbind()),
        PropertyValue::Time(v) => {
            // Windows FILETIME (100-nanosecond intervals since January 1, 1601)
            // to Unix timestamp (seconds since January 1, 1970)
            let windows_epoch = 116444736000000000u64; // January 1, 1601 in 100-nanosecond intervals
            let unix_timestamp = (*v as u64).saturating_sub(windows_epoch) / 10_000_000;
            let datetime_module = py.import("datetime")?;
            let datetime_class = datetime_module.getattr("datetime")?;
            let datetime =
                datetime_class.call_method1("fromtimestamp", (unix_timestamp as i64,))?;
            Ok(datetime.unbind())
        }
        PropertyValue::Guid(v) => {
            // GUIDを文字列に変換
            let guid_str = format!("{:?}", v);
            Ok(guid_str.into_pyobject(py)?.into_any().unbind())
        }
        PropertyValue::Binary(v) => Ok(PyBytes::new(py, v.buffer()).into_any().unbind()),
        PropertyValue::Object(_) => {
            // Objectは複雑なので、とりあえずNoneを返す
            Ok(py.None())
        }
        PropertyValue::MultipleInteger16(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::MultipleInteger32(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::MultipleFloating32(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::MultipleFloating64(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::MultipleCurrency(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::MultipleFloatingTime(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::MultipleInteger64(v) => Ok(v.into_pyobject(py)?.into_any().unbind()),
        PropertyValue::MultipleString8(v) => {
            let strings: Vec<String> = v.iter().map(|s| s.to_string()).collect();
            Ok(strings.into_pyobject(py)?.into_any().unbind())
        }
        PropertyValue::MultipleUnicode(v) => {
            let strings: Vec<String> = v.iter().map(|s| s.to_string()).collect();
            Ok(strings.into_pyobject(py)?.into_any().unbind())
        }
        PropertyValue::MultipleTime(v) => {
            let windows_epoch = 116444736000000000u64;
            let timestamps: Vec<i64> = v
                .iter()
                .map(|&t| ((t as u64).saturating_sub(windows_epoch) / 10_000_000) as i64)
                .collect();
            let datetime_module = py.import("datetime")?;
            let datetime_class = datetime_module.getattr("datetime")?;
            let datetimes: Vec<PyObject> = timestamps
                .iter()
                .map(|&ts| {
                    datetime_class
                        .call_method1("fromtimestamp", (ts,))
                        .map(|obj| obj.unbind())
                })
                .collect::<PyResult<Vec<_>>>()?;
            Ok(datetimes.into_pyobject(py)?.into_any().unbind())
        }
        PropertyValue::MultipleGuid(v) => {
            let guid_strings: Vec<String> = v.iter().map(|g| format!("{:?}", g)).collect();
            Ok(guid_strings.into_pyobject(py)?.into_any().unbind())
        }
        PropertyValue::MultipleBinary(v) => {
            let binaries: Vec<PyObject> = v
                .iter()
                .map(|b| PyBytes::new(py, b.buffer()).into_any().unbind())
                .collect();
            Ok(binaries.into_pyobject(py)?.into_any().unbind())
        }
    }
}

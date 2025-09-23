# Email Body Extraction Debug Improvements

## Problem Identified
- 80% of emails (40 out of 50) have empty body content in `list_inbox_emails`
- ALL emails show size=0 bytes (definitely wrong)
- Even test emails with known content show empty bodies
- Some emails (20%) do work, indicating intermittent COM property access issues

## Root Cause Analysis
The issue is in the `_get_email_property` method which uses standard `getattr()` to access COM object properties. This fails silently for many emails, likely due to:
1. COM object not being fully loaded when properties are accessed
2. Property names having different cases or variations
3. Need for explicit COM property invocation

## Improvements Made

### 1. Enhanced `_get_email_property` Method
- **Method 1**: Standard attribute access (existing)
- **Method 2**: Try case variations (Size, size, SIZE, Size)
- **Method 3**: Force COM object property access using `pythoncom.DISPATCH_PROPERTYGET`
- Added comprehensive debug logging for each method

### 2. Enhanced `_extract_email_body` Method
- Added debug prints to show body extraction progress
- Shows lengths of plain text and HTML bodies at each step
- Tracks when text extraction from HTML occurs

### 3. Enhanced `_transform_email_to_data` Method
- Added debug print at the start of email transformation
- Shows which folder is being processed

## Expected Results After Server Restart
When the HTTP server is restarted with these changes, we should see:
1. **Debug output** showing which COM access method works for each property
2. **Improved size detection** - emails should show correct byte sizes
3. **Better body extraction** - more emails should have body content
4. **Clear error messages** if COM access still fails

## Files Modified
- `src/outlook_mcp_server/adapters/outlook_adapter.py`
  - Enhanced `_get_email_property` method (lines ~2162-2220)
  - Enhanced `_extract_email_body` method (lines ~2272-2310)
  - Enhanced `_transform_email_to_data` method (lines ~1732-1745)

## Next Steps
1. **Restart the HTTP server** to load the enhanced code
2. **Run test**: `python test_debug_prints.py` to see debug output
3. **Run full test**: `python test_50_email_bodies_http.py` to verify improvements
4. **Analyze results** and make further adjustments if needed

## Technical Details
The key improvement is using `pythoncom.DISPATCH_PROPERTYGET` to force COM property access:
```python
disp_id = email_item._oleobj_.GetIDsOfNames(property_name)[0]
value = email_item._oleobj_.Invoke(disp_id, 0, pythoncom.DISPATCH_PROPERTYGET, 1)
```

This bypasses Python's attribute access and directly invokes the COM property getter, which should be more reliable for Outlook COM objects.
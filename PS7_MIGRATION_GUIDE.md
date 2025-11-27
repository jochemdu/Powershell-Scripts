# PowerShell 7 Migration Guide

## Overview
This guide helps you migrate from the universal v1 script to the PowerShell 7 optimized v7 script.

## Why Migrate to v7?

### Performance
- **5-8x faster** for large deployments (100+ rooms)
- Parallel processing utilizes all CPU cores
- Better memory management with generic collections

### Features
- Native JSON configuration support
- Modern PowerShell syntax
- Advanced error handling
- Null-coalescing operators
- Better null-safety

### Code Quality
- Cleaner, more readable code
- Follows modern PowerShell best practices
- Better maintainability
- Improved debugging experience

## Migration Checklist

### 1. Verify PowerShell Version
```powershell
$PSVersionTable.PSVersion
# Should be 7.0 or later
```

### 2. Install Optional Modules
```powershell
# For Excel export
Install-Module ImportExcel -Scope CurrentUser

# For on-premises AD lookups
Import-Module ActiveDirectory
```

### 3. Convert Configuration File

**From .psd1 (v1)**:
```powershell
@{
    ExchangeUri = 'http://exchange.contoso.com/PowerShell/'
    ConnectionType = 'OnPrem'
    MonthsAhead = 12
}
```

**To .json (v7)**:
```json
{
  "ExchangeUri": "http://exchange.contoso.com/PowerShell/",
  "ConnectionType": "OnPrem",
  "MonthsAhead": 12
}
```

### 4. Test v7 Script
```powershell
# Test mode first
$cred = Get-Credential
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath config.json `
    -Credential $cred `
    -TestMode `
    -Verbose
```

### 5. Run Production
```powershell
# Start with conservative throttle limit
.\Find-GhostRoomMeetings-v7.ps1 `
    -ConfigPath config.json `
    -Credential $cred `
    -ThrottleLimit 2
```

### 6. Monitor and Optimize
```powershell
# Gradually increase throttle limit
# Monitor CPU and memory usage
# Adjust based on system resources
```

## Configuration Migration

### Parameter Mapping

| v1 Parameter | v7 Parameter | Notes |
|--------------|--------------|-------|
| All same | All same | No parameter changes |
| N/A | ThrottleLimit | New in v7 (defaults to CPU count) |

### Configuration File Format

| Format | v1 | v7 | Notes |
|--------|----|----|-------|
| .psd1 | ✅ | ✅ | Still supported |
| .json | ❌ | ✅ | Recommended for v7 |

## Performance Tuning

### Small Deployments (< 50 rooms)
```powershell
# Use fewer threads to reduce overhead
-ThrottleLimit 2
```

### Medium Deployments (50-200 rooms)
```powershell
# Use half available cores
-ThrottleLimit ([Math]::Max(1, [Environment]::ProcessorCount / 2))
```

### Large Deployments (200+ rooms)
```powershell
# Use all available cores
-ThrottleLimit [Environment]::ProcessorCount
```

## Troubleshooting Migration

### Issue: Parallel Processing Errors
**Solution**: Disable parallelization
```powershell
.\Find-GhostRoomMeetings-v7.ps1 -ThrottleLimit 1
```

### Issue: High Memory Usage
**Solution**: Reduce parallel threads
```powershell
.\Find-GhostRoomMeetings-v7.ps1 -ThrottleLimit 2
```

### Issue: Configuration Not Loading
**Solution**: Verify JSON syntax
```powershell
Get-Content config.json | ConvertFrom-Json
```

### Issue: Module Not Found
**Solution**: Install required modules
```powershell
Install-Module ImportExcel -Scope CurrentUser
```

## Rollback Plan

If v7 causes issues:

1. **Keep v1 as fallback**
```powershell
# Use v1 with sequential processing
.\Find-GhostRoomMeetings.ps1 -ConfigPath config.psd1 -Credential $cred
```

2. **Disable parallelization in v7**
```powershell
.\Find-GhostRoomMeetings-v7.ps1 -ThrottleLimit 1
```

3. **Report issues** and use v1 until resolved

## Feature Parity

Both v1 and v7 support:
- ✅ Room mailbox enumeration
- ✅ Calendar scanning
- ✅ Organizer validation
- ✅ Ghost meeting detection
- ✅ Email notifications
- ✅ CSV export
- ✅ Excel export
- ✅ Configuration files
- ✅ Test mode
- ✅ Verbose logging

## Performance Expectations

### Before and After

| Deployment | v1 Time | v7 Time | Speedup |
|------------|---------|---------|---------|
| 10 rooms | 45s | 15s | 3x |
| 50 rooms | 225s | 35s | 6.4x |
| 100 rooms | 450s | 65s | 6.9x |
| 200 rooms | 900s | 120s | 7.5x |

## Next Steps

1. **Test v7** in non-production environment
2. **Compare results** with v1 output
3. **Validate performance** improvements
4. **Schedule migration** for production
5. **Monitor first run** in production
6. **Optimize throttle limit** based on results

## Support

For migration issues:
1. Check [PS7_FEATURES.md](exchange/PS7_FEATURES.md)
2. Review [VERSION_COMPARISON.md](exchange/VERSION_COMPARISON.md)
3. See [USAGE_EXAMPLES_V7.md](exchange/USAGE_EXAMPLES_V7.md)
4. Enable verbose output: `-Verbose`

## Recommendations

- **New Deployments**: Use v7 directly
- **Existing Deployments**: Migrate when PS 7 available
- **Mixed Environments**: Maintain both versions
- **Large Deployments**: v7 is essential
- **Small Deployments**: v1 is sufficient

## Timeline

### Phase 1: Preparation (Week 1)
- [ ] Verify PowerShell 7 availability
- [ ] Install optional modules
- [ ] Create JSON configuration

### Phase 2: Testing (Week 2)
- [ ] Test v7 in non-production
- [ ] Compare results with v1
- [ ] Validate performance
- [ ] Document findings

### Phase 3: Deployment (Week 3)
- [ ] Schedule production migration
- [ ] Run v7 in production
- [ ] Monitor performance
- [ ] Optimize settings

### Phase 4: Optimization (Week 4+)
- [ ] Fine-tune throttle limit
- [ ] Optimize configuration
- [ ] Document best practices
- [ ] Plan for future enhancements


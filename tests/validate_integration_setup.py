#!/usr/bin/env python3
"""
Validation script for integration test setup.

This script checks if the system is properly configured for running
integration tests with the Outlook MCP Server.
"""

import sys
import os
from pathlib import Path
from typing import Dict, List, Tuple

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))


def check_python_version() -> Tuple[bool, str]:
    """Check if Python version is compatible."""
    version = sys.version_info
    if version.major == 3 and version.minor >= 8:
        return True, f"Python {version.major}.{version.minor}.{version.micro} ✓"
    else:
        return False, f"Python {version.major}.{version.minor}.{version.micro} ✗ (requires 3.8+)"


def check_required_modules() -> Tuple[bool, List[str]]:
    """Check if all required modules are available."""
    required_modules = [
        ("win32com.client", "Windows COM interface"),
        ("pythoncom", "Python COM support"),
        ("pytest", "Testing framework"),
        ("asyncio", "Async support"),
        ("psutil", "System monitoring"),
        ("json", "JSON support"),
        ("datetime", "Date/time support"),
        ("threading", "Threading support"),
        ("concurrent.futures", "Concurrent execution"),
        ("tempfile", "Temporary files"),
        ("pathlib", "Path handling")
    ]
    
    results = []
    all_available = True
    
    for module_name, description in required_modules:
        try:
            __import__(module_name)
            results.append(f"  {module_name}: ✓ Available")
        except ImportError:
            results.append(f"  {module_name}: ✗ Missing ({description})")
            all_available = False
    
    return all_available, results


def check_outlook_availability() -> Tuple[bool, str]:
    """Check if Microsoft Outlook is available."""
    try:
        from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
        
        adapter = OutlookAdapter()
        if adapter.connect():
            # Test basic operations
            try:
                namespace = adapter.get_namespace()
                inbox = adapter.get_folder_by_name("Inbox")
                adapter.disconnect()
                return True, "Microsoft Outlook: ✓ Available and accessible"
            except Exception as e:
                adapter.disconnect()
                return False, f"Microsoft Outlook: ✗ Connected but not fully accessible ({str(e)})"
        else:
            return False, "Microsoft Outlook: ✗ Cannot connect (not running or not installed)"
            
    except ImportError as e:
        return False, f"Microsoft Outlook: ✗ Import error ({str(e)})"
    except Exception as e:
        return False, f"Microsoft Outlook: ✗ Connection error ({str(e)})"


def check_system_resources() -> Tuple[bool, List[str]]:
    """Check system resources and configuration."""
    results = []
    all_good = True
    
    try:
        import psutil
        
        # Check available memory
        memory = psutil.virtual_memory()
        memory_gb = memory.total / (1024**3)
        if memory_gb >= 4:
            results.append(f"  Memory: ✓ {memory_gb:.1f} GB available")
        else:
            results.append(f"  Memory: ⚠ {memory_gb:.1f} GB (recommend 4+ GB)")
            all_good = False
        
        # Check CPU
        cpu_count = psutil.cpu_count()
        results.append(f"  CPU Cores: ✓ {cpu_count} cores")
        
        # Check disk space
        disk = psutil.disk_usage('.')
        disk_gb = disk.free / (1024**3)
        if disk_gb >= 1:
            results.append(f"  Disk Space: ✓ {disk_gb:.1f} GB free")
        else:
            results.append(f"  Disk Space: ⚠ {disk_gb:.1f} GB free (recommend 1+ GB)")
            all_good = False
            
    except ImportError:
        results.append("  System Resources: ⚠ Cannot check (psutil not available)")
        all_good = False
    except Exception as e:
        results.append(f"  System Resources: ✗ Error checking ({str(e)})")
        all_good = False
    
    return all_good, results


def check_test_files() -> Tuple[bool, List[str]]:
    """Check if test files are present and accessible."""
    test_files = [
        "test_integration.py",
        "integration_test_config.py", 
        "run_integration_tests.py",
        "pytest_integration.ini",
        "INTEGRATION_TESTING.md"
    ]
    
    results = []
    all_present = True
    test_dir = Path(__file__).parent
    
    for test_file in test_files:
        file_path = test_dir / test_file
        if file_path.exists():
            results.append(f"  {test_file}: ✓ Present")
        else:
            results.append(f"  {test_file}: ✗ Missing")
            all_present = False
    
    return all_present, results


def check_source_files() -> Tuple[bool, List[str]]:
    """Check if source files are present."""
    src_dir = Path(__file__).parent.parent / "src" / "outlook_mcp_server"
    
    required_files = [
        "server.py",
        "main.py",
        "adapters/outlook_adapter.py",
        "services/email_service.py",
        "services/folder_service.py",
        "protocol/mcp_protocol_handler.py",
        "routing/request_router.py",
        "error_handler.py",
        "logging/logger.py",
        "models/mcp_models.py"
    ]
    
    results = []
    all_present = True
    
    for src_file in required_files:
        file_path = src_dir / src_file
        if file_path.exists():
            results.append(f"  {src_file}: ✓ Present")
        else:
            results.append(f"  {src_file}: ✗ Missing")
            all_present = False
    
    return all_present, results


def check_permissions() -> Tuple[bool, str]:
    """Check if we have necessary permissions."""
    try:
        # Test file creation in current directory
        test_file = Path("test_permissions.tmp")
        test_file.write_text("test")
        test_file.unlink()
        
        # Test temp directory access
        import tempfile
        with tempfile.NamedTemporaryFile() as tmp:
            tmp.write(b"test")
        
        return True, "File Permissions: ✓ Can create files and directories"
        
    except Exception as e:
        return False, f"File Permissions: ✗ Cannot create files ({str(e)})"


def generate_recommendations(checks: Dict[str, Tuple[bool, any]]) -> List[str]:
    """Generate recommendations based on check results."""
    recommendations = []
    
    # Python version
    if not checks["python"][0]:
        recommendations.append("Upgrade to Python 3.8 or higher")
    
    # Missing modules
    if not checks["modules"][0]:
        missing_modules = [
            line.split(":")[0].strip() for line in checks["modules"][1] 
            if "✗" in line
        ]
        if "win32com.client" in str(missing_modules):
            recommendations.append("Install pywin32: pip install pywin32")
        if "pytest" in str(missing_modules):
            recommendations.append("Install pytest: pip install pytest pytest-asyncio")
        if "psutil" in str(missing_modules):
            recommendations.append("Install psutil: pip install psutil")
    
    # Outlook
    if not checks["outlook"][0]:
        recommendations.append("Start Microsoft Outlook and ensure it's properly configured")
        recommendations.append("Verify Outlook is not in offline mode")
        recommendations.append("Consider running as administrator if COM access fails")
    
    # System resources
    if not checks["resources"][0]:
        recommendations.append("Close unnecessary applications to free up system resources")
        recommendations.append("Consider running tests on a system with more resources")
    
    # Test files
    if not checks["test_files"][0]:
        recommendations.append("Ensure all integration test files are present in the tests directory")
    
    # Source files
    if not checks["source_files"][0]:
        recommendations.append("Ensure the Outlook MCP Server source code is complete")
    
    # Permissions
    if not checks["permissions"][0]:
        recommendations.append("Run with administrator privileges or fix file permissions")
    
    return recommendations


def main():
    """Main validation function."""
    print("Outlook MCP Server Integration Test Setup Validation")
    print("=" * 60)
    
    # Run all checks
    checks = {
        "python": check_python_version(),
        "modules": check_required_modules(),
        "outlook": check_outlook_availability(),
        "resources": check_system_resources(),
        "test_files": check_test_files(),
        "source_files": check_source_files(),
        "permissions": check_permissions()
    }
    
    # Display results
    print("\n1. Python Version:")
    print(f"   {checks['python'][1]}")
    
    print("\n2. Required Modules:")
    for line in checks["modules"][1]:
        print(line)
    
    print(f"\n3. Microsoft Outlook:")
    print(f"   {checks['outlook'][1]}")
    
    print("\n4. System Resources:")
    for line in checks["resources"][1]:
        print(line)
    
    print("\n5. Test Files:")
    for line in checks["test_files"][1]:
        print(line)
    
    print("\n6. Source Files:")
    for line in checks["source_files"][1]:
        print(line)
    
    print(f"\n7. Permissions:")
    print(f"   {checks['permissions'][1]}")
    
    # Overall status
    all_passed = all(check[0] for check in checks.values())
    
    print("\n" + "=" * 60)
    if all_passed:
        print("✓ ALL CHECKS PASSED - Ready for integration testing!")
    else:
        print("✗ SOME CHECKS FAILED - See recommendations below")
        
        recommendations = generate_recommendations(checks)
        if recommendations:
            print("\nRecommendations:")
            for i, rec in enumerate(recommendations, 1):
                print(f"  {i}. {rec}")
    
    print("\nNext Steps:")
    if all_passed:
        print("  1. Run quick test: python tests/run_integration_tests.py --quick")
        print("  2. Run full suite: python tests/run_integration_tests.py --verbose")
        print("  3. Generate report: python tests/run_integration_tests.py --report")
    else:
        print("  1. Address the issues listed in recommendations")
        print("  2. Re-run this validation script")
        print("  3. Proceed with integration testing once all checks pass")
    
    print("=" * 60)
    
    return 0 if all_passed else 1


if __name__ == "__main__":
    sys.exit(main())
#!/usr/bin/env python3
"""
Unified E2E Test Suite Runner for "04 - GOOD CARD" (Cencosud)
Discovers and executes all test tiers (Tier 1 Features, Tier 2 Boundaries,
Tier 3 Combinations, Tier 4 Workloads) and returns exit code 0 on complete pass.
"""

import os
import sys
import time
import unittest

# Ensure project root is in sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)


def run_e2e_suite() -> bool:
    print("=" * 80)
    print(" CENCOSUD - AUDITORIA 04 - GOOD CARD : E2E TEST SUITE RUNNER")
    print("=" * 80)
    print(f"Project Root: {PROJECT_ROOT}")
    print(f"Python:       {sys.version.split()[0]}")
    print(f"Platform:     {sys.platform}")
    print("-" * 80)

    start_time = time.time()

    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    # Discover tests from tests/e2e directory
    e2e_dir = os.path.dirname(__file__)
    discovered = loader.discover(start_dir=e2e_dir, pattern="test_tier*.py")
    suite.addTests(discovered)

    total_discovered = suite.countTestCases()
    print(f"Discovered: {total_discovered} test cases across all E2E tiers.\n")

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    duration = time.time() - start_time

    print("\n" + "=" * 80)
    print(" TEST EXECUTION SUMMARY")
    print("=" * 80)
    print(f"Total Tests Run:  {result.testsRun}")
    print(f"Passed:           {result.testsRun - len(result.failures) - len(result.errors) - len(result.skipped)}")
    print(f"Failures:         {len(result.failures)}")
    print(f"Errors:           {len(result.errors)}")
    print(f"Skipped:          {len(result.skipped)}")
    print(f"Execution Time:   {duration:.3f} seconds")
    print("-" * 80)

    if result.wasSuccessful():
        print(">>> RESULT: ALL E2E TEST TIERS PASSED SUCCESSFULLY (EXIT CODE 0) <<<")
        print("=" * 80)
        return True
    else:
        print(">>> RESULT: TEST FAILURES DETECTED <<<")
        print("=" * 80)
        return False


if __name__ == "__main__":
    success = run_e2e_suite()
    sys.exit(0 if success else 1)

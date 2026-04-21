#!/usr/bin/env python3
"""
Simple performance test to verify optimization improvements.
Run this to measure operation speeds.
"""
import time
import inventory_engine as engine

def measure_time(func, *args, label=""):
    """Measure execution time of a function."""
    start = time.time()
    result = func(*args)
    elapsed = (time.time() - start) * 1000  # Convert to milliseconds
    print(f"⏱️  {label}: {elapsed:.1f}ms")
    return result, elapsed

def run_performance_tests():
    """Run performance benchmarks."""
    print("\n" + "="*50)
    print("📊 PERFORMANCE TEST SUITE")
    print("="*50 + "\n")
    
    # Test 1: Database Migration
    print("Test 1: Database Migration (schema check)")
    measure_time(engine.migrate_db, label="migrate_db()")
    
    # Test 2: Get Engineers
    print("\nTest 2: Read Engineer Data")
    result, time_ms = measure_time(engine.get_engineers, label="get_engineers()")
    print(f"   Loaded {len(result)} engineers")
    
    # Test 3: Get Service Jobs
    print("\nTest 3: Read Service Jobs")
    result, time_ms = measure_time(engine.get_service_jobs, label="get_service_jobs()")
    if result is not None:
        print(f"   Loaded {len(result)} service jobs")
    
    # Test 4: Get Articles
    print("\nTest 4: Read Master Articles")
    def get_articles():
        df = engine._read_sheet_df("Master", engine.MASTER_COLUMNS)
        return df
    result, time_ms = measure_time(get_articles, label="read_master_sheet()")
    if result is not None and not result.empty:
        print(f"   Loaded {len(result)} articles")
    
    print("\n" + "="*50)
    print("✅ Performance tests completed!")
    print("="*50 + "\n")
    print("💡 Tips for best performance:")
    print("   • Use 'Reformat Excel' button periodically to maintain speed")
    print("   • Excel formatting is skipped during normal operations")
    print("   • Large datasets benefit from regular database backups/resets")
    print()

if __name__ == "__main__":
    run_performance_tests()

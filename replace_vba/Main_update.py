from config import config
from s1_copysource import copy_and_process_data
from s2_compareRanges import compare_ranges
from s3_addDeleteCodes import add_delete_code
from s4_combineAllsource import combine_all_source
from s5_importWork import import_work
from s8_copyCoscoActual import copy_cosco_actual

def run_update_process():
    print("Starting update process...\n")
    
    execution_flow = [
        ("Step 1: Copying source data...", copy_and_process_data),
        ("Step 2: Comparing ranges...", compare_ranges),
        ("Step 3: Adding and deleting codes...", add_delete_code),
        ("Step 4: Combining all source data...", combine_all_source),
        ("Step 5: Importing work data...", import_work),
        ("Step 6: Copying COSCO actual data...", copy_cosco_actual),
    ]

    for step_msg, processor in execution_flow:
        try:
            print(step_msg)
            processor()
            print(f"✅ {step_msg} completed.\n")
        except Exception as e:
            print(f"❌ Critical failure in {step_msg}: {str(e)}")
            raise SystemExit(1)

    print("🎉 Success: All pipeline stages executed")
    print(f"Final config state: {config.__dict__}")

if __name__ == "__main__":
    run_update_process()
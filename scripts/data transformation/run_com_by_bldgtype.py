#Run Com.py once per building-type workbook in DEER_EnergyPlus_Modelkit_Measure_list_working_dir,
#stashing each pass's outputs under outputs_by_bldgtype/<BldgType>/.
#Each pass copies the per-type workbook onto the plain workbook name Com.py reads,
#and sets COM_BT_FILTER so Com.py only reads that building type's runs.
#Afterwards the plain workbook is restored to the Asm copy (the committed version).
import os, shutil, subprocess, sys, time

os.chdir(os.path.dirname(os.path.abspath(__file__)))

WB_DIR = "DEER_EnergyPlus_Modelkit_Measure_list_working_dir"
PLAIN = "DEER_EnergyPlus_Modelkit_Measure_list_working.xlsx"
OUT = "outputs_by_bldgtype"
OUTPUTS = [
    "current_msr_mat.csv",
    "sim_annual.csv",
    "sim_hourly_wb.csv",
    "CEDARS_LoadShape_Com.zip",
    "CEDARS_ls_annual_loads_Com.csv",
]
# CEDARS_long_ls_Com.csv duplicates the zip content uncompressed (~hundreds of MB
# per building type); deleted after each pass instead of stashed.
DISCARD = ["CEDARS_long_ls_Com.csv"]

wbs = sorted(f for f in os.listdir(WB_DIR) if f.endswith(".xlsx") and not f.startswith("~"))
print(f"{len(wbs)} workbooks to process", flush=True)

failures = []
for wb in wbs:
    bt = wb.rsplit(" ", 1)[-1][:-5]  # '... OFC EPr.xlsx' -> 'EPr'
    dest = os.path.join(OUT, bt)
    os.makedirs(dest, exist_ok=True)
    shutil.copyfile(os.path.join(WB_DIR, wb), PLAIN)
    env = dict(os.environ, COM_BT_FILTER=bt)
    t0 = time.time()
    with open(os.path.join(dest, "com_run.log"), "w") as log:
        r = subprocess.run([sys.executable, "Com.py"], stdout=log, stderr=subprocess.STDOUT, env=env)
    moved = []
    for f in OUTPUTS:
        if os.path.exists(f):
            shutil.move(f, os.path.join(dest, f))
            moved.append(f)
    for f in DISCARD:
        if os.path.exists(f):
            os.remove(f)
    status = "OK" if r.returncode == 0 else f"FAIL rc={r.returncode}"
    if r.returncode != 0:
        failures.append(bt)
    print(f"{bt}: {status}, {len(moved)}/{len(OUTPUTS)} outputs, {time.time()-t0:.0f}s", flush=True)

# leave the plain workbook as the Asm (committed) version
shutil.copyfile(os.path.join(WB_DIR, "DEER_EnergyPlus_Modelkit_Measure_list_working OFC Asm.xlsx"), PLAIN)

print(f"done. failures: {failures if failures else 'none'}", flush=True)
sys.exit(1 if failures else 0)

import pandas as pd, json

with open(f"settings/config.json", "r", encoding="utf-8") as file:
    config = json.load(file)

# 무조건 name - email 순이어야함.
headers = config["target_cols"]

df = pd.read_csv("settings/data.csv", usecols=headers)

# variable
fail_list = []
when = None

while True:
    if not when:
        when = str(input("When (mm-dd): "))
        if not when:
            if input("exit? "): continue
            else: break

    email = str(input("Email: "))
    if not email:
        when = None
        continue

    # copy 안붙이면 에러남
    target = df[df[headers[0]] == email].copy()
    if not target.empty:
        row = target.iloc[0]
        print(f"{row[headers[0]]}: {row[headers[1]]} found!")
        # reason check
        while True:
            reason = str(input("Reason: "))
            if reason:
                target["사유"] = reason
                break
        target["일자"] = when
        fail_list.append(target)
    else: print("mail not found")

if fail_list:
    final_df = pd.concat(fail_list, ignore_index=True)
    final_df.to_csv("settings/fail_data.csv", mode="w", index=False, encoding="utf-8")
    print("save complete")

input("press enter to exit")
from __future__ import annotations

from collections import Counter
from datetime import date
import json
from pathlib import Path
from statistics import median

import pandas as pd


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / "source-data" / "分贝通明细全年_2026-07-23_1618.xlsx"
JULY_SOURCE = ROOT / "july_records.json"
QUOTA_SOURCE = ROOT / "quota_records.json"
MONTHS = [f"{month}月" for month in range(1, 8)]
TRAVEL_TYPES = {"国内机票": "机票", "火车": "火车", "酒店": "酒店"}


def number(value) -> float:
    if pd.isna(value):
        return 0.0
    return float(value)


def text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def feishu_value(value):
    if isinstance(value, list):
        return "".join(
            str(item.get("text", "")) if isinstance(item, dict) else str(item)
            for item in value
        )
    if isinstance(value, dict) and "value" in value:
        return value["value"]
    return value


def records_frame(path: Path) -> pd.DataFrame:
    records = json.loads(path.read_text(encoding="utf-8"))
    return pd.DataFrame(
        [
            {key: feishu_value(value) for key, value in record["fields"].items()}
            for record in records
        ]
    )


def group_from_department(value: str) -> str:
    department = text(value)
    for group in ("客户成功1组", "客户成功2组", "客户成功3组", "客户成功4组"):
        if group in department:
            return group
    return "部门直属"


def top_items(counter: Counter, limit: int) -> list[dict]:
    return [
        {"name": name, "value": round(value, 2)}
        for name, value in counter.most_common(limit)
    ]


detail = pd.read_excel(SOURCE, sheet_name="客户成功部分贝通明细")
detail = pd.concat([detail, records_frame(JULY_SOURCE)], ignore_index=True)
quota = records_frame(QUOTA_SOURCE)

detail = detail[
    detail["月份"].isin(MONTHS) & detail["业务线"].isin(TRAVEL_TYPES)
].copy()
detail["amount"] = pd.to_numeric(
    detail["企业支付总金额"], errors="coerce"
).fillna(0)
detail["name"] = detail["预订|下单|用餐人"].map(text)
detail["department"] = detail["预订|下单人层级部门"].map(text)
detail["group"] = detail["department"].map(group_from_department)

quota = quota[quota["部门"].astype(str).str.contains("客户成功部", na=False)].copy()
quota_by_name = {text(row["姓名"]): row for _, row in quota.iterrows()}
all_people = sorted(set(detail["name"]) | set(quota_by_name))

people = []
for name in all_people:
    rows = detail[detail["name"] == name]
    quota_row = quota_by_name.get(name)
    latest_rows = rows[rows["月份"] == MONTHS[-1]]
    if not latest_rows.empty:
        department = text(latest_rows["department"].mode().iloc[0])
    elif not rows.empty:
        department = text(rows["department"].mode().iloc[0])
    elif quota_row is not None:
        department = text(quota_row["部门"])
    else:
        department = ""
    group = group_from_department(department)

    q2_quota = number(quota_row.get("二季度额度（5-7月）", 0)) if quota_row is not None else 0
    monthly = []
    for month_index, month in enumerate(MONTHS, 1):
        month_rows = rows[rows["月份"] == month]
        spent = float(month_rows["amount"].sum())
        month_quota = (
            number(quota_row.get(f"{month_index}月差旅", 0))
            if quota_row is not None and month_index <= 4
            else q2_quota / 3
        )
        monthly.append(
            {
                "month": month,
                "spent": round(spent, 2),
                "quota": round(month_quota, 2),
                "usage": spent / month_quota if month_quota else None,
                "records": int((month_rows["amount"] > 0).sum()),
            }
        )

    spent = float(rows["amount"].sum())
    period_quota = sum(item["quota"] for item in monthly)
    usage = spent / period_quota if period_quota else None
    positive_rows = rows[rows["amount"] > 0]
    active_months = int(positive_rows["月份"].nunique())

    cities = Counter()
    routes = Counter()
    types = Counter()
    refund_records = 0
    for _, row in rows.iterrows():
        amount = number(row["amount"])
        origin = text(row["出发城市|下单城市|取车城市"])
        destination = text(row["到达城市|入住城市|目的城市|用餐城市|还车城市"])
        status = text(row["机票|订单|支付状态"])
        expense_type = TRAVEL_TYPES[text(row["业务线"])]
        types[expense_type] += amount
        if destination and amount > 0:
            cities[destination] += amount
        if origin and destination and amount > 0:
            routes[f"{origin}-{destination}"] += amount
        if ("退" in status or amount < 0):
            refund_records += 1

    people.append(
        {
            "name": name,
            "department": department,
            "group": group,
            "spent": round(spent, 2),
            "quota": round(period_quota, 2),
            "usage": usage,
            "remaining": round(period_quota - spent, 2),
            "records": int((rows["amount"] > 0).sum()),
            "activeMonths": active_months,
            "cityCount": len(cities),
            "refundRecords": refund_records,
            "topCities": top_items(cities, 6),
            "topRoutes": top_items(routes, 4),
            "byType": {kind: round(types.get(kind, 0), 2) for kind in ("机票", "火车", "酒店")},
            "monthly": monthly,
            "labels": [],
            "riskScore": 0,
        }
    )

group_medians = {}
for group in sorted({person["group"] for person in people}):
    amounts = [
        person["spent"]
        for person in people
        if person["group"] == group and person["spent"] > 0
    ]
    group_medians[group] = median(amounts) if amounts else 0

for person in people:
    labels = []
    score = 0
    if person["quota"] == 0 and person["spent"] > 0:
        labels.append("有费用但未匹配额度")
        score += 6
    elif person["usage"] is not None and person["usage"] > 1:
        labels.append("超过个人限制额度")
        score += 5
    elif person["usage"] is not None and person["usage"] >= 0.8:
        labels.append("接近个人限制额度")
        score += 3

    latest_month_spent = person["monthly"][-1]["spent"]
    historical = [item["spent"] for item in person["monthly"][:-1] if item["spent"] > 0]
    historical_average = sum(historical) / len(historical) if historical else 0
    if latest_month_spent >= 3000 and historical_average and latest_month_spent >= historical_average * 1.8:
        labels.append("7月费用显著上升")
        score += 2

    peer_median = group_medians[person["group"]]
    if person["spent"] >= 5000 and peer_median and person["spent"] >= peer_median * 1.8:
        labels.append("高于同组中位数")
        score += 2
    if person["refundRecords"] >= 3:
        labels.append("退款/退票记录较多")
        score += 1
    if person["cityCount"] >= 6:
        labels.append("出行城市较多")
        score += 1

    person["labels"] = labels or ["暂无异常"]
    person["riskScore"] = score
    person["peerMedian"] = round(peer_median, 2)
    person["risk"] = "高" if score >= 5 else "中" if score >= 3 else "低"

monthly_summary = []
for month_index, month in enumerate(MONTHS):
    month_rows = detail[detail["月份"] == month]
    spent = float(month_rows["amount"].sum())
    month_quota = sum(person["monthly"][month_index]["quota"] for person in people)
    monthly_summary.append(
        {
            "month": month,
            "spent": round(spent, 2),
            "quota": round(month_quota, 2),
            "usage": spent / month_quota if month_quota else None,
            "records": int((month_rows["amount"] > 0).sum()),
            "people": int(month_rows.loc[month_rows["amount"] > 0, "name"].nunique()),
            "cities": int(
                month_rows.loc[
                    month_rows["amount"] > 0,
                    "到达城市|入住城市|目的城市|用餐城市|还车城市",
                ].nunique()
            ),
        }
    )

groups = []
for group in sorted({person["group"] for person in people}):
    members = [person for person in people if person["group"] == group]
    spent = sum(person["spent"] for person in members)
    group_quota = sum(person["quota"] for person in members)
    groups.append(
        {
            "group": group,
            "people": len(members),
            "expensePeople": sum(person["spent"] != 0 for person in members),
            "spent": round(spent, 2),
            "quota": round(group_quota, 2),
            "usage": spent / group_quota if group_quota else None,
            "remaining": round(group_quota - spent, 2),
            "records": sum(person["records"] for person in members),
            "cities": len(
                {
                    city["name"]
                    for person in members
                    for city in person["topCities"]
                }
            ),
            "warningPeople": sum(person["riskScore"] >= 3 for person in members),
            "overPeople": sum(
                person["usage"] is not None and person["usage"] > 1
                for person in members
            ),
            "nearPeople": sum(
                person["usage"] is not None and 0.8 <= person["usage"] <= 1
                for person in members
            ),
        }
    )

total_spent = sum(person["spent"] for person in people)
total_quota = sum(person["quota"] for person in people)
overrun_total = sum(max(0, -person["remaining"]) for person in people)
unused_total = sum(max(0, person["remaining"]) for person in people)
people_by_spend = sorted(people, key=lambda person: person["spent"], reverse=True)
type_totals = {
    kind: round(sum(person["byType"][kind] for person in people), 2)
    for kind in ("机票", "火车", "酒店")
}
all_cities = {
    city["name"]
    for person in people
    for city in person["topCities"]
}

data = {
    "generatedAt": date.today().isoformat(),
    "period": "2026年1–7月",
    "notes": {
        "expense": "企业支付净额，包含退款、退票等负向冲减",
        "quota": "1–4月使用月度额度；5–7月按5–7月季度额度÷3折算月均额度",
        "coverage": "城市与路线用于描述交付出行覆盖，不代表交付结果或质量",
    },
    "summary": {
        "people": len(people),
        "expensePeople": sum(person["spent"] != 0 for person in people),
        "spent": round(total_spent, 2),
        "quota": round(total_quota, 2),
        "usage": total_spent / total_quota if total_quota else None,
        "remaining": round(total_quota - total_spent, 2),
        "records": sum(person["records"] for person in people),
        "cities": len(all_cities),
        "warningPeople": sum(person["riskScore"] >= 3 for person in people),
        "nearPeople": sum(
            person["usage"] is not None and 0.8 <= person["usage"] <= 1
            for person in people
        ),
        "overPeople": sum(
            person["usage"] is not None and person["usage"] > 1
            for person in people
        ),
        "overrunTotal": round(overrun_total, 2),
        "unusedTotal": round(unused_total, 2),
        "top5Share": (
            sum(person["spent"] for person in people_by_spend[:5]) / total_spent
            if total_spent
            else 0
        ),
        "top10Share": (
            sum(person["spent"] for person in people_by_spend[:10]) / total_spent
            if total_spent
            else 0
        ),
        "typeTotals": type_totals,
    },
    "monthly": monthly_summary,
    "groups": groups,
    "people": people,
}

(ROOT / "data.json").write_text(
    json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
)
print(
    json.dumps(
        {
            "people": data["summary"]["people"],
            "spent": data["summary"]["spent"],
            "quota": data["summary"]["quota"],
            "usage": data["summary"]["usage"],
            "groups": len(groups),
        },
        ensure_ascii=False,
    )
)

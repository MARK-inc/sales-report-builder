# generates a messy sample export (fake data) to demo the cleaner
import csv, random, datetime as dt
random.seed(4)
products = {"Espresso beans 1kg": 89, "Grinder burr set": 240, "Milk jug 600ml": 55, "Pour-over kit": 160, "Descaler 4-pack": 39}
fmts = ["%d/%m/%Y", "%Y-%m-%d", "%d %b %Y", "%m-%d-%Y"]
rows, oid = [], 1000
for i in range(420):
    d = dt.date(2026, 1, 1) + dt.timedelta(days=random.randint(0, 272))
    p = random.choice(list(products)); q = random.choice([1, 1, 1, 2, 2, 3])
    amt = products[p] * q
    a = random.choice([f"₪{amt:,.2f}", f"{amt}", f"{amt:.1f}".replace(".", ","), f"ILS {amt}"])
    f = random.choice(fmts)
    if f == "%m-%d-%Y": f = "%Y-%m-%d"  # keep it unambiguous
    name = random.choice(["  Noa Levi", "Dan Cohen ", "AVI MIZRAHI", "maya katz", "Lior Ben-David", "Shira Peretz"])
    rows.append([f"#{oid}", d.strftime(f), name, p.upper() if random.random() < .1 else p, q, a])
    oid += 1
# noise: duplicates, empty amounts, bad dates
for r in random.sample(rows, 18): rows.append(list(r))
for r in random.sample(rows, 6): r[5] = ""
for r in random.sample(rows, 4): r[1] = "31/02/2026"
random.shuffle(rows)
with open("orders_export.csv", "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh); w.writerow(["Order ID", "Date", "Customer", "Product", "Qty", "Total"]); w.writerows(rows)
print(len(rows), "rows written")

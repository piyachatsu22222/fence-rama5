# -*- coding: utf-8 -*-
"""
วิเคราะห์ PO งานรั้ว โครงการ 4111 — ตัวเรียกต่อ (shim)

⚠️ โค้ดจริงย้ายไปรวมกับสคริปต์หลักแล้ว: สรุปยอด_PO_รั้วพระราม5.py --mode wbs
ไฟล์นี้เก็บไว้เพื่อให้คำสั่ง/ทางลัดเดิมยังใช้ได้ ไม่ต้องแก้อะไร

ใช้แบบเดิม (ยังได้อยู่):
    py -3 วิเคราะห์_PO_รั้ว_4111.py [ไฟล์ 4111_po.xlsx] [ไฟล์ผลลัพธ์.txt]

ใช้แบบใหม่ (แนะนำ — มีตัวเลือกครบกว่า):
    py -3 สรุปยอด_PO_รั้วพระราม5.py --mode wbs --source "...\4111_po.xlsx" --out สรุป.txt
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from importlib import import_module

หลัก = import_module("สรุปยอด_PO_รั้วพระราม5")


def main() -> int:
    ที่มา = sys.argv[1] if len(sys.argv) > 1 else None
    ผลลัพธ์ = sys.argv[2] if len(sys.argv) > 2 else None

    if not ที่มา:
        print(__doc__)
        print("ต้องระบุไฟล์ต้นทาง เช่น:")
        print('    py -3 วิเคราะห์_PO_รั้ว_4111.py "C:\\...\\4111_po.xlsx"')
        return 1

    อาร์กิวเมนต์ = ["--mode", "wbs", "--source", ที่มา]
    if ผลลัพธ์:
        อาร์กิวเมนต์ += ["--out", ผลลัพธ์]
    return หลัก.main(อาร์กิวเมนต์)


if __name__ == "__main__":
    sys.exit(main())

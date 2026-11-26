# THAMC Inventory — ระบบจัดการคลังเวชภัณฑ์

เว็บแอปไฟล์เดียว (`index.html`) บน GitHub Pages + ฐานข้อมูล Google Sheet ผ่าน Apps Script

## 1. เตรียม Google Sheet
แท็บที่ต้องมี: `FRZ294` (คอลัมน์ A–M ตามไฟล์ล่าสุด) และ `Sheet3` (รหัสกลุ่ม / ชื่อกลุ่ม)
แท็บ `Extra` (REF, บาร์โค้ด, รูป 1–3, New!), `Locations` (ตู้/ชั้น/ช่อง) และ `Log` จะถูกสร้างโดย `setup()`

## 2. Deploy Apps Script
1. เปิด Sheet → Extensions → Apps Script → วางโค้ดจาก `apps-script/Code.gs`
2. เลือกฟังก์ชัน `setup` → Run → อนุญาตสิทธิ์ (Sheets + Drive)
3. ตั้ง `ADMIN_PIN` ใน `CFG` (ค่าเริ่มต้น 1234 ต้องเปลี่ยน) และ (ถ้าต้องการ) ใส่ `TOKEN` ใน `CFG` เป็นรหัสผ่านร่วม และ `PHOTO_FOLDER_ID` ของโฟลเดอร์รูป
4. Deploy → New deployment → Web app
   - Execute as: **Me**
   - Who has access: **Anyone** (หรือ Anyone with Google account ถ้าไม่ใช้ TOKEN)
5. คัดลอก URL ที่ลงท้าย `/exec`
6. ทุกครั้งที่แก้โค้ด ต้อง Deploy → Manage deployments → New version

ทดสอบ: เปิด `<URL>/exec?action=ping` ต้องได้ `{"ok":true,...}`

## 3. Deploy GitHub Pages
1. สร้าง repo (เช่น `thamc-inventory`) แล้วอัปโหลดไฟล์ `index.html` ไว้ที่ root
2. Settings → Pages → Source: `main` / root
3. เปิด `https://<user>.github.io/thamc-inventory/`

## 4. เชื่อมต่อ
- เปิดเว็บ → แถบซ้ายล่าง กด **ตั้งค่า API** → วาง URL `/exec` (และรหัสผ่านถ้าตั้ง TOKEN)
  ค่าจะเก็บในเครื่อง (localStorage) แจกลิงก์ให้เครื่อง handheld ได้แบบ `…/index.html?api=<URL>`

## การทำงาน
| ฟังก์ชัน | เก็บที่ |
|---|---|
| Location | `FRZ294` คอลัมน์ D + บันทึกประวัติใน `Log` |
| เคลื่อนไหว / ไม่เคลื่อนไหว | `FRZ294` คอลัมน์ K |
| REF, บาร์โค้ดผู้ผลิต (หลายค่า) | `Extra` (คั่นด้วย `|`) |
| รูป 1–3 | Google Drive (ลิงก์ใน `Extra`, รูปแรกใส่ `FRZ294` คอลัมน์ J ด้วย) |

- QR บนป้ายมีเนื้อหา = **Item code** (เช่น `6510630002`); Barcode 1D = Code 128
- สแกนด้วยกล้อง: ใช้ BarcodeDetector (Chrome / Android); เครื่อง handheld ที่มีปุ่มสแกนในตัวให้ยิงเข้าช่อง "พิมพ์ / ยิงบาร์โค้ด"
- เปิดตรงที่ไอเทม: `index.html#/item/6510630002` หรือ `?item=6510630002`
- สแกนหา: Item code, TUH Code, บาร์โค้ดผู้ผลิต หรือ REF

## สิทธิ์
- ผู้ใช้ทั่วไป: ดู ค้นหา สแกน และพิมพ์ป้ายได้ แก้ไขไม่ได้
- แอดมิน: กดปุ่ม "แอดมิน" มุมซ้ายล่างแล้วใส่ `ADMIN_PIN` จึงแก้ Location / REF / บาร์โค้ด / รูป / สถานะ และนำเข้าข้อมูลได้
- ฝั่งเซิร์ฟเวอร์ตรวจ PIN ทุกครั้งที่เขียนข้อมูล (ไม่ใช่แค่ซ่อนปุ่มในหน้าเว็บ)

# ويندوز سيرفر أوديت 🖥️

أداة لفحص خوادم ويندوز عن بُعد — جمع بيانات المستخدمين المحليين، المجموعات المحلية، والمهام المجدولة — مع إمكانية تغيير كلمة مرور المدير المحلي. يتم الاتصال عبر WinRM، وتُشفّر كلمات المرور باستخدام Fernet، والمخرجات تكون تقارير Excel منسّقة مع سجل أحداث (.log).

---

## الملفات الأساسية (3 ملفات فقط)

| الملف | الوظيفة |
|---|---|
| `src/crypto.py` | إنشاء مفتاح التشفير (`--init`) وتشفير كلمات المرور (`--set`) |
| `src/audit.py` | فحص الخوادم (قراءة فقط) — المستخدمين، المجموعات، المهام |
| `src/change_password.py` | تغيير كلمة مرور المدير المحلي على جميع الخوادم |

---

## الإعداد الأولي

### 1. البيئة الافتراضية والمتطلبات

```powershell
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

### 2. إنشاء مفتاح التشفير

```powershell
python src/crypto.py --init
```

ينشئ الملف `config/secret.key`

### 3. إدخال كلمات المرور وتشفيرها

```powershell
python src/crypto.py --set
```

يُطلب منك إدخال كلمة مرور أو أكثر (واحدة تلو الأخرى). يتم تشفيرها وحفظها في `config/passwords.enc`.

### 4. تحضير قائمة الخوادم

أنشئ ملف `config/servers.txt` — اسم خادم أو IP في كل سطر. الأسطر الفارغة والتي تبدأ بـ `#` يتم تجاهلها.

```
SRV-APP01
192.168.1.50
# SRV-DB01  (مؤقتاً خارج الخدمة)
SRV-WEB02
```

---

## تشغيل الفحص (قراءة فقط)

```powershell
python src/audit.py
```

**المخرجات:**
- `output/Audit_Report_YYYYMMDD_HHMMSS.xlsx` — تقرير Excel يحتوي على الأوراق:
  - **Summary** — ملخص لكل خادم (عدد المستخدمين، المشرفين، المجموعات، المهام)
  - **Server Login Status** — حالة الاتصال بكل خادم
  - **Local Users and Privileges** — كل المستخدمين مع تصنيف الصلاحيات
  - **Local Groups and Members** — كل المجموعات والأعضاء
  - **Scheduled Tasks** — المهام المجدولة
  - **Errors** — الأخطاء (في حال وجودها)
- `output/Audit_YYYYMMDD_HHMMSS.log` — سجل الأحداث كامل مع الوقت

---

## تشغيل تغيير كلمة المرور

```powershell
python src/change_password.py
```

يُطلب منك إدخال كلمة المرور الجديدة مرتين (تأكيد). يتم تغيير كلمة مرور المدير المحلي على كل الخوادم.

**المخرجات:**
- `output/Password_Change_YYYYMMDD_HHMMSS.xlsx` — تقرير بحالة التغيير
- `output/Password_Change_YYYYMMDD_HHMMSS.log` — سجل الأحداث

---

## استخدام ملفات الإعداد من أي مكان على القرص

يمكنك وضع ملفات الإعداد (`servers.txt`، `secret.key`، `passwords.enc`) في **أي مجلد** على القرص الصلب، ثم تحديد مسارها عند التشغيل:

```powershell
# جميع الملفات في مجلد مختلف
python src/audit.py --servers "D:\Team\config\myservers.txt" --key "D:\Team\keys\secret.key" --passwords "D:\Team\keys\passwords.enc" --output "D:\Team\reports"

# استخدام المجلد الافتراضي للإخراج فقط مع ملفات إعداد من مكان آخر
python src/change_password.py --servers "C:\Users\Admin\Desktop\servers.txt" --key "E:\secrets\secret.key" --passwords "E:\secrets\passwords.enc"
```

### الخيارات المتاحة:

| الخيار | القيمة الافتراضية | الوصف |
|---|---|---|
| `--servers` | `config/servers.txt` | مسار ملف قائمة الخوادم |
| `--key` | `config/secret.key` | مسار مفتاح التشفير |
| `--passwords` | `config/passwords.enc` | مسار ملف كلمات المرور المشفّرة |
| `--output` | `output` | مجلد إخراج التقارير |

**ملاحظة:** `--output` يمكن أن يكون مساراً كاملاً أيضاً:
```powershell
python src/audit.py --output "D:\Team\Audit Results\July 2026"
```

---

## مثال كامل (سير عمل متكامل)

```powershell
# 1. تفعيل البيئة
.\venv\Scripts\activate

# 2. إنشاء مفتاح التشفير
python src/crypto.py --init

# 3. إدخال كلمات المرور
python src/crypto.py --set

# 4. تشغيل الفحص على جميع الخوادم
python src/audit.py

# 5. مراجعة التقرير في المجلد output/

# 6. تغيير كلمة المرور
python src/change_password.py
```

---

## بنية المجلدات

```
windows-server-audit/
├── config/                     # ملفات الإعداد الافتراضية
│   ├── servers.txt             # قائمة الخوادم
│   ├── secret.key              # مفتاح التشفير (مستبعد من git)
│   └── passwords.enc           # كلمات المرور المشفّرة (مستبعد من git)
├── output/                     # التقارير وسجلات الأحداث (مستبعد من git)
│   ├── Audit_Report_*.xlsx
│   ├── Audit_*.log
│   ├── Password_Change_*.xlsx
│   └── Password_Change_*.log
├── src/
│   ├── crypto.py               # إدارة التشفير
│   ├── audit.py                # أداة الفحص
│   └── change_password.py      # أداة تغيير كلمة المرور
├── doc/                        # وثائق تفصيلية (إنجليزية)
├── requirements.txt
└── README.ar.md                # هذا الملف
```

---

## الأمان

- `secret.key` و `passwords.enc` مستبعدان من git (موجودان في `.gitignore`) — لا ترفعهم أبداً إلى المستودع.
- `audit.py` للقراءة فقط — لا ينفذ أي أوامر تعديل (`Set-LocalUser`، `Add-LocalGroupMember`، إلخ).
- `change_password.py` يعدّل **فقط** حساب Administrator المحلي ولا يمس أي شيء آخر.
- جميع الأخطاء تُسجّل لكل عملية على حدة — فشل خادم واحد لا يوقف الفحص بالكامل.

---

## المتطلبات

- Python 3.10+
- خوادم ويندوز مفعّلة عليها WinRM و CredSSP
- حساب المشرف المحلي على الخوادم المستهدفة

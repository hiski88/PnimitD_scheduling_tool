# מע׳ לתכנון תורנויות- פנימית ד׳

אפליקציית Streamlit עבור מע׳ לתכנון תורנויות- פנימית ד׳.


## תפריט צד

תחת הכותרת **כלי עזר** מופיעים שלושה כלים:

1. **תכנון תורנויות** — ברירת המחדל, הכלי שבנוי כרגע להזנת זמינות, חופשות וחסימות.
2. **קידוד לטבלה כללית** — הדבקת פלטים מכל העובדים, יצירת טבלת חודש עם עמודות ידניות ועובדים, סימון חסימות באפור והורדת Excel.
3. **שמירה ביומן** — placeholder לכלי עתידי לשמירת תורנויות ליומן.


## מה יש בדמו

- ברירת מחדל: הצגת החודש הבא.
- מעבר חופשי בין חודשים ללא מגבלה.
- טבלת תאריכים עם:
  - שבתות וחגים יהודיים/ישראליים כתצוגה בלבד.
  - אירועים מיומן אישי כתצוגה בלבד.
  - סימון ידני של חסימה לתורנות.
  - סימון ידני של יום חופש.
  - העדפה רכה והערה.
- כפתור "סיום / שמור" עם פלט טקסטואלי להעתקה.
- יצירת פלט גנרי מסוג `PersonAvailability`.
- שמירה לקובץ `data/availability_submissions.jsonl`.
- הורדת JSON אחרי שמירה.
- כלי קידוד לטבלה כללית עם הורדת Excel.

## התקנה מקומית

```bash
pip install -r requirements.txt
streamlit run app.py
```

## פריסה ל-Streamlit Cloud

1. העלה את הקבצים ל-GitHub.
2. צור אפליקציה ב-Streamlit Community Cloud.
3. הגדר Secrets אם רוצים Google Calendar.
4. ודא שה-redirect URI ב-Google Cloud זהה לכתובת האפליקציה.

## Google Calendar

האפליקציה משתמשת בהרשאת קריאה בלבד:

```text
https://www.googleapis.com/auth/calendar.readonly
```

### הגדרת Google Cloud

1. צור Project ב-Google Cloud.
2. Enable ל-Google Calendar API.
3. צור OAuth Client מסוג Web application.
4. הוסף Authorized redirect URI, לדוגמה:

```text
https://YOUR_APP.streamlit.app
```

5. הוסף ב-Streamlit Secrets:

```toml
[google_oauth]
client_id = "YOUR_CLIENT_ID"
client_secret = "YOUR_CLIENT_SECRET"
redirect_uri = "https://YOUR_APP.streamlit.app"
```

## עקרון תכנוני

אירוע ביומן או חג אינו הופך אוטומטית לאילוץ.
המשתמש רואה מידע, בוחר ידנית, ורק אז המערכת מייצרת אילוצים למנוע השיבוץ.

מבנה פלט לדוגמה:

```json
{
  "person_id": "demo_user_001",
  "date": "2026-06-12",
  "type": "unavailable_for_shift",
  "strength": "hard",
  "source": "manual",
  "note": ""
}
```

## שלבים עתידיים

- שמירה ל-PostgreSQL במקום JSONL.
- אימות משתמשים.
- תמיכה ב-Outlook Calendar.
- הרשאות לפי ארגון/מחלקה.
- הפרדת Holiday Provider לשירות עצמאי.
- שילוב עם מנוע השיבוץ.

// server.js

const express = require('express');
// נייבא את הפונקציה הראשית שלנו מהקובץ השני
const { generateScheduleFromData } = require('./schedule-generator.js');
const cors = require('cors');
const morgan = require('morgan');
const app = express();
// Middleware שמאפשר לשרת לקרוא גוף בקשה בפורמט JSON
app.use(express.json());

const PORT = process.env.PORT || 3330;


app.use(cors()); // מאפשר גישה מכל מקור
app.use(morgan('dev')); // Middleware להדפסת בקשות לשרת בקונ
// הגדרת נקודת קצה (Endpoint) ליצירת הסידור
// אנחנו משתמשים ב-POST כי הלקוח (Postman) שולח לנו מידע (dbData)
app.post('/generate-schedule', async (req, res) => {
    console.log('התקבלה בקשה ליצירת סידור עבודה...');

    // המידע שהגיע מהלקוח (Postman) נמצא ב-req.body
    const dbData = req.body;

    // בדיקה בסיסית שהמידע הגיע
    if (!dbData || !dbData.schedule) {
        return res.status(400).json({ message: "Bad Request: Missing schedule data in request body." });
    }

    try {
        // קריאה לפונקציה המרכזית שלנו עם המידע שהתקבל
        await generateScheduleFromData(dbData);
        
        // אם הכל עבר בהצלחה, שלח תגובת הצלחה
        res.status(200).json({ message: '✅ הטבלה הדינמית נוצרה בהצלחה!' });

    } catch (err) {
        // אם קרתה שגיאה כלשהי בתהליך, נתפוס אותה ונשלח תגובת שגיאה
        console.error('❌ אירעה שגיאה בתהליך:', err.message);
        res.status(500).json({ 
            message: '❌ אירעה שגיאה בתהליך', 
            error: err.message 
        });
    }
});

app.get('/', (req, res) => {
    res.send('ברוכים הבאים לשרת יצירת לוחות זמנים! השתמשו בנקודת הקצה /generate-schedule ליצירת סידור עבודה.');
});

app.listen(PORT, () => {
    console.log(`שרת מאזין בפורט ${PORT}`);
});
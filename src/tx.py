import datetime
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, Set, Optional


class Category(Enum):
    mortgage = "משכנתא"
    food = "אוכל"
    shvut = "שבות"
    education = "חינוך"
    running_expenses = "שוטף"
    mentoring = "הדרכה"
    donation = "תרומה"
    tax = "מס"
    insurance = "ביטוח"
    atm = "כספומט"
    fuel = "דלק"
    savings = "חסכון"
    transport = "תחבורה"
    other = "אחר"
    income = "הכנסות"


descriptions_by_category: Dict[Category, Set[str]] = {
    Category.transport: {
        "WIND MOBILITY",
        "רב קו",
        "LIM*RIDE",
    },
    Category.mortgage: {
        "משכנתא",
    },
    Category.tax: {
        "מסים",
        "מיסים",
    },
    Category.running_expenses: {
        "ועד",
        "סלקום",
        "חברת החשמל",
        "019",
        "פלאפון",
        "בזק",
        "אלקטרה מוצרי צריכה",
        "אמישרגז",
        "מי אביבים"
    },
    Category.savings: {
        "חסכון",
    },
    Category.donation: {
        "פעמוני",
        "חיים ביד",
        "מוסדות חב\"ד",
        "מכון מאיר",
        "עטרת",
        "חסדי עולם",
        "מה יפו פעמי",
        "מאירים ביפו",
        "התורה והארץ",
        "ישיבת הג",  # הגולן
        "גרעין יפו",
        "בית דוד בית שמש",
        "המרכז העולמי לחסד",
        "אסתר המלכה",
        "יד שרה",
        "ויצמן שולה",
        "עין-דרור",
    },
    Category.insurance: {
        "מכבי",
        "שירותי ברי",
        "ביטוח",
        "כלל ב.בריאות",
        "מבטחים",
        "פניקס",
        "מגדל",
    },
    Category.education: {
        "אמונה",
        "אורות התורה",
        'ישיבת בנ"ע',
        'עיריית בת ים',
        "חינוך",
    },
    Category.atm: {
        "כספומט"
    },
    Category.mentoring: {
        "שר שלום",
        "שפר",
    },
    Category.fuel: {
        "פנגו",
        "דור אלון",
        "דלק",
        "סונול",
        "פז",
        "מיקה",
        "כלל חובה",
        "כלל אלמנטרי",
        "כלל רכב",
        "רישיונות רכב",
        "חניוני",
        "משרד התחבורה",
        "רכב דוד",
    },
    Category.shvut: {
        "מאיה אלגריסי",
        "צאלה קרני",
        "דורית אילני",
        "קרני צאלה",
    },
    Category.food: {
        "מכולת",
        "יוחננוף",
        "סופר דוש",
        "קפה עלית",
        "יסמין עייש",
        "מגה קמעונאות",
        "יינות ביתן",
        "טוב זה בטבע",
        "ויקטורי",
        "חצות וחצי",
        "מגה בעיר",
        "מחסני להב",
        "שופרסל",
        "פיצה",
        "אושר עד",
        "רמי לוי",
        "מינימרקט",
        "מוצרי מזון",
        "סיבוס",
    },
}

category_by_description: Dict[str, Category] = {
    keyword: cat for cat, keywords in descriptions_by_category.items() for keyword in keywords
}


@dataclass
class Transaction:
    amount: Any
    business: str
    charge_date: datetime.datetime
    transaction_date: datetime.datetime
    details: str
    card: str
    notes: str
    transaction_sum: Any
    tid: Optional[str] = None

    def __post_init__(self) -> None:
        self.category = self._compute_category()

    def _compute_category(self) -> Category:
        if self.amount < 0:
            return Category.income
        for kw, category in category_by_description.items():
            if kw in self.business:
                return category
        return Category.other

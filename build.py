#!/usr/bin/env python3
"""Build the spending dashboard.

Usage:  python3 build.py path/to/transactions.xlsx [out.html]

Reads a transaction export (Rocket Money / Copilot style columns), cleans and
de-duplicates it, classifies every transaction into a category of our own,
normalises vendor names, aggregates everything the dashboard needs and writes
a single self-contained HTML file (template.html + embedded JSON).
"""
import json
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent

# --------------------------------------------------------------------------
# Taxonomy: category -> group.  Groups are what the stacked charts use (<= 8).
# --------------------------------------------------------------------------
GROUPS = {
    "Housing": ["Rent & Housing"],
    "Transportation": ["Car Payments", "Gas & Convenience", "Auto Service & Parts",
                       "Car Wash", "Tolls & Parking", "Rideshare & Rentals",
                       "Vehicle Purchase & Fees"],
    "Food & Dining": ["Restaurants", "Fast Food", "Coffee & Bakery",
                      "Food Delivery", "Groceries"],
    "Shopping": ["General Merchandise", "Online Shopping", "Clothing & Accessories",
                 "Electronics & Gaming", "Home & Furniture", "Books & Hobbies"],
    "Bills & Insurance": ["Utilities & Internet", "Phone", "Insurance"],
    "Subscriptions & Tech": ["Streaming & Media", "Software & AI Tools",
                             "Music Production"],
    "Lifestyle & Travel": ["Entertainment & Events", "Travel", "Personal Care",
                           "Health & Fitness", "Medical", "Education",
                           "Business & Side Hustle"],
    "Money & People": ["Church & Charity", "Family Support", "Payments to People",
                       "Cash Withdrawals", "Fees & Interest", "Credit & Debt Services",
                       "Buy Now Pay Later", "Other"],
}
CAT2GROUP = {c: g for g, cs in GROUPS.items() for c in cs}

ESSENTIAL = {"Rent & Housing", "Car Payments", "Gas & Convenience", "Auto Service & Parts",
             "Tolls & Parking", "Groceries", "Utilities & Internet", "Phone", "Insurance",
             "Medical", "Vehicle Purchase & Fees", "Education"}

# Non-spending buckets (never counted as spend)
X_CC = "Credit Card Payment"
X_TRANSFER = "Transfer"
X_INVEST = "Investing"
X_INCOME = "Income"
X_REWARD = "Rewards & Adjustments"
EXCLUDED = {X_CC, X_TRANSFER, X_INVEST, X_INCOME, X_REWARD}

# --------------------------------------------------------------------------
# Rules: (regex, vendor-or-None, category).  First match wins.  Matched
# against UPPER("name | description").  Vendor None => cleaned name.
# --------------------------------------------------------------------------
R = [
    # ---- non-spending ------------------------------------------------------
    (r"CREDIT CARD PAYMENT|PAYMENT - THANK YOU|MOBILE PAYMENT|MOB PAYMENT RECEIVED|NFO PAYMENT RECEIVED"
     r"|^PAYMENT RECEIVED|AMEX EPAYMENT|CREDIT ONE BANK PAYMENT|ACI\*CREDIT ONE|PAYMENT - MOBILE APP"
     r"|PAYMENT - DEBIT CARD|NAVY FEDERAL CC|^CAPITAL ONE \||AUTOPAY PAYMENT|ONLINE PAYMENT", None, X_CC),
    (r"REWARDS? CREDIT|CASH REWARDS|ATM REBATE|SECURITY ADJUSTMENT|ACCOUNT ADJUSTMENT|CREDIT PROTECTION ADJ"
     r"|FINANCE CHARGE ADJ|ATM ADJUSTMENT|OD FEE WINDOW|RETURNED CHECK|REWARD POINTS|INTEREST ADJ", None, X_REWARD),
    (r"CRYPTO\.COM ARENA", "Crypto.com Arena", "Entertainment & Events"),
    (r"CHURCH'?S CHICKEN|CHURCHS CHICKEN|CHURCH S CHICKEN", "Church's Chicken", "Fast Food"),
    (r"CYMATICS", "Cymatics", "Music Production"),
    (r"APRIVA|365 MARKET|MARKET J\b", "365 Market (vending)", "Coffee & Bakery"),
    (r"FID BKG SVC|MONEYLINE|ROBINHOOD|MOONPAY|CRO ST JULIANS|CRYPTO\.COM|COINBASE|FIDELITY INVESTMENTS"
     r"|CASH APP\*KING DIAMO|WEBULL|ACORNS|STASH", None, X_INVEST),
    (r"DIRECT DEP|PAYROLL|TWC-BENEFITS|TWC BENEFITS|DEPOSIT@MOBILE|ATM DEPOSIT|ETSY INC DEPOSIT|ETSY PAYOUT"
     r"|SHOPIFY TRANSFER|SEZZLE PAYOUT|CTU REFUND|INTEREST PAID|IRS TREAS|TAX REF", None, X_INCOME),
    (r"CASH APP\*JAYSHUN|CASH APP\*CASH OUT|JAYSHUN SEPHUS|SEPHUS JAYSHUN|UPHO\*J ?SEPHUS|UPHO\*JAYSHUN"
     r"|INTERNET DEPOSIT|INTERNET WITHDRAWAL|ACCTVERIFY|PENNY TEST|ALLY BANK ACCT FUND|TRANSFER TO|TRANSFER FROM"
     r"|OPTUM BANK|^VENMO|CHASE\d+ W PARKER|SIMPLE XFERS", None, X_TRANSFER),

    # ---- housing -----------------------------------------------------------
    (r"COOL SPRINGS", "Cool Springs Apartments", "Rent & Housing"),
    (r"CENTRAL SQUARE", "Central Square Apartments", "Rent & Housing"),
    (r"MERITAGE AT STEINER", "Meritage at Steiner Ranch", "Rent & Housing"),
    (r"BELL-B\d|BELL PARTNERS", "Bell Apartments", "Rent & Housing"),
    (r"LEBANON RIDGE", "Lebanon Ridge Apartments", "Rent & Housing"),
    (r"HUNTER WARFIELD", "Hunter Warfield (apartment collections)", "Rent & Housing"),
    (r"PEAK AUTO STORAGE|PUBLIC STORAGE|EXTRA SPACE|CUBESMART", None, "Rent & Housing"),

    # ---- transportation ----------------------------------------------------
    (r"ALLY (ALLY )?PAYMT|ALLY RETRY", "Ally Auto", "Car Payments"),
    (r"CAPITAL ONE AUTO", "Capital One Auto", "Car Payments"),
    (r"GM FINANCIAL", "GM Financial", "Car Payments"),
    (r"MEPCO|ENDURANCE", "Endurance Warranty", "Car Payments"),
    (r"GARLYN SHELTON|CARMAX|AUTOTRADER|MCKINNEY BUICK|CARVANA|VEHICLE REG|TXDMV|TX DMV", None, "Vehicle Purchase & Fees"),
    (r"BMW|INTEGRITY-1ST|AUTOZONE|O'?REILLY|ADVANCE AUTO|ROCK ?AUTO|BUDGET WRENCH|SERVICE STREET|AAA TIRE"
     r"|FIRESTONE|DISCOUNT TIRE|JIFFY LUBE|VALVOLINE|TAKE 5|PEP BOYS|MIDAS|MEINEKE|WILLIE`?S PERFORMANCE"
     r"|AMERICAN DREAM AUTO|MAVIS|GOODYEAR|BUMPER\.COM|AUTO PARTS|TIRE|COLLISION|AUTOMOTIVE|MOTORS\b", None, "Auto Service & Parts"),
    (r"CAR ?WASH|CARWASH|QWIKWASH|WHIP MY SOUL|TODAYS CAR", None, "Car Wash"),
    (r"NTTA|TOLL|PARKING|PARKMOBILE|PARKWHIZ|PARK N FLY|PARK 'N FLY|STREET METERS|PARKING METERS|PACPARK"
     r"|PLAT PARKING|SP\+|LAZ PARK|ABM PARK", None, "Tolls & Parking"),
    (r"UBER ?\*? ?EATS|UBEREATS|DOORDASH|DD \*DOORDASH|GRUBHUB|POSTMATES|INSTACART.*REST|FAVOR DELIVERY", None, "Food Delivery"),
    (r"\bUBER\b|UBR\*|\bLYFT\b|TURO|HERTZ|ENTERPRISE RENT|\bAVIS\b|U-HAUL|UHAUL|BIRD APP|LIM\*RIDE"
     r"|CHARLIE CAR RENTAL|BUDGET CAR RENTAL|ZIPCAR|DART |GETAROUND", None, "Rideshare & Rentals"),
    (r"SHELL|CHEVRON|EXXON|MOBIL\b|RACETRAC|RACE TRAC|\bQT\b|QUIKTRIP|CEFCO|CIRCLE ?K|7-ELEVEN|7ELEVEN|7 ELEVEN"
     r"|STAR MART|BUC-EE|DIAMOND SHAMROCK|SUNOCO|TEXACO|LOVE ?S TRAVEL|VALERO|PHILLIPS 66|CONOCO|MURPHY"
     r"|SPEEDWAY|MAVERIK|WAWA|CORNER STORE|QUICK CHECK|POWER MART|GATEWAY FOOD MART|LUCKY FOOD MART"
     r"|CORNER MARKET|FOOD STORE|FOOD MART|STOP N GO|STRIPES|\bFUEL\b|\bGAS\b|PMUSA|C STORE|LEGACY MART|MINI MART"
     r"|TEXAN MART|ALS SUPER CENTER", None, "Gas & Convenience"),

    # ---- bills & insurance -------------------------------------------------
    (r"GEICO|PROGRESSIVE|PROG COUNTY|USAA INSURANCE|NOBLR|JETTY INS|RENTERS|CONDO INS|HOMEOWNERS INS|ALLIANZ"
     r"|STATE FARM|ALLSTATE|LEMONADE|UNITED OF OMAHA|VISION SERVICE PLAN|GUARDIAN DENTAL|INSURANCE|\bINS PREM"
     r"|TEXAS LAWSHIELD|U\.?S\.? LAWSHIELD", None, "Insurance"),
    (r"VERIZON|VZWRLSS|AT&T|ATT\*|ATT BILL|\bATT\b|T-MOBILE|TMOBILE|SPECTRUM MOBILE|CRICKET|BOOST MOBILE"
     r"|METRO BY|MINT MOBILE|PINGER|VZ ?WIRELESS|VZW", None, "Phone"),

    # ---- money & people ----------------------------------------------------
    (r"INTEREST CHARGE|INTEREST ON PURCHASES|LATE FEE|LATE PAYMENT|CREDIT PROTECT|NSF FEE|OVERDRAFT|OD FEE"
     r"|MONTHLY SERVICE FEE|ATM FEE|NON-CHASE ATM FEE|ANNUAL FEE|RENEWAL MEMBERSHIP FEE|EXPRESS PAYMENT FEE"
     r"|FOREIGN TRANSACTION|SERVICE CHARGE|\bFEE\b", None, "Fees & Interest"),
    (r"EXPERIAN|TRANSUNION|TU \*|EQUIFAX|LEXINGTON LAW|CREDIT KARMA|FED DEBT|ROCKET MONEY|TRUEBILL"
     r"|DEBT PAYOFF|TIMELYBILLS|CHANGED INC|CHANGED? INC", None, "Credit & Debt Services"),
    (r"AFFIRM|KLARNA|SEZZLE|AFTERPAY|ZIP\* |ZIP\*|QUADPAY|PAYPAL PAY IN 4", None, "Buy Now Pay Later"),
    (r"GODS WAY|GOD'S WAY|CHURCH|MINISTR|TITHE|ALFREDO PA|CASH APP\*TIFF|T\.D\. JAKES|POTTERS HOUSE|OFFERING"
     r"|GIVELIFY|TITHE\.LY|PUSHPAY|RED CROSS|ST JUDE|GOFUNDME|CHARIT|DONAT", None, "Church & Charity"),
    (r"RONDA SEPH|ROY SEPHUS|SEPHUS", None, "Family Support"),
    (r"ATM WITHDRAW|ATM WITHDRAWAL|WITHDRAWAL-USAA|NON-CHASE ATM WITHDRAW|^KOHL'S \||TRANSFUND|PAI ISO|EFT \d+ ATM"
     r"|CASH WITHDRAWAL", None, "Cash Withdrawals"),
    (r"^CASH APP|^ZELLE|PAYPAL VISA DIRECT|APPLE CASH|CASH APP\*", None, "Payments to People"),

    # ---- subscriptions & tech ---------------------------------------------
    (r"FLUXTATION|SSAIYAD|SERATO|SPLICE|SOUNDSTRIPE|TRACKLIB|MIXUNIT|THE MUSIC UNIT|AVID TECH|NATIVE INSTRUMENTS"
     r"|PLUGIN BOUTIQUE|SWEETWATER|GUITAR CENTER|BEATSTARS|DISTROKID|LANDR|WAVES AUDIO|ABLETON|IMAGE-LINE"
     r"|LOOPMASTERS|SPINSTER RECORDS|SOUNDCLOUD|AMI MUSICBOX", None, "Music Production"),
    (r"NETFLIX HOUSE", "Netflix House", "Entertainment & Events"),
    (r"NETFLIX|HULU|HLU\*|SPOTIFY|STARZ|SHOWTIME|PARAMOUNT|CBS |CBS\b|PEACOCK|DISNEY|HBO|MAX\.COM|HELP\.MAX|YOUTUBE"
     r"|AUDIBLE|TIDAL|APPLE\.COM/BILL|APPLE MUSIC|ROKU|PATREON|XBOX|PLAYSTATION|NINTENDO|ESPN|NFL|DALLAS MORNING NEWS"
     r"|WSJ|WALL-ST|DOW JONES|D J\*|BUSINESS JOURNALS|BNP MEDIA|SCENTBIRD|SBD\*|BROWN SUGAR|CRUNCHYROLL|SLING"
     r"|FUBO|PHILO|AMAZON PRIME|PRIME VIDEO|KINDLE|SIRIUS|PANDORA|CHESS\.COM|CHESS COM|CLASSMATES|MOMENTHOUSE"
     r"|TWITTER|X CORP|INSTAGRAM|MEET NOW", None, "Streaming & Media"),
    (r"REPLIT|CLAUDE|ANTHROPIC|OPENAI|CHATGPT|PERPLEXITY|PERPLEXIT|NOTION|ADOBE|MICROSOFT|MSBILL|GOOGLE ONE|GOOGLE CLOUD"
     r"|GOOGLE WORKSPACE|GOOGLE \*|GOOGLE STORAGE|TRADINGVIEW|CANVA|BEFUNKY|TODOIST|LUCID|PLANOLY|SQUARESPACE"
     r"|GODADDY|WIX|SHOPIFY|DROPBOX|ICLOUD|LOVABLE|CURSOR|GITHUB|VISILY|APIFY|RAPIDAPI|POLYGON|MASSIVE\.COM"
     r"|DIVTRACKER|LEMSQZY|PDF|EARTHLINK|GOTO GROUP|XME INC|PICSART|PLANNER5D|PLANNER 5D|PLACEIT|MEGA LIMITED"
     r"|DTSI|ATTSERVICESINC|RESUME-NOW|RESUME NOW|CAPCUT|ELEVENLABS|MIDJOURNEY|ZOOM|SOFTWARE|\.AI\b|APP STORE", None,
     "Software & AI Tools"),

    # ---- lifestyle & travel -----------------------------------------------
    (r"HILTON|HAMPTON INN|DOUBLETREE|MARRIOTT|HYATT|HOLIDAY INN|FOUR POINTS|SHERATON|WESTIN|AIRBNB|VRBO|EXPEDIA"
     r"|HOTELS\.COM|BOOKING\.COM|HOPPER|AIRLINES|AIR LINES|DELTA AIR|UNITED \d|\bUNITED\b|SOUTHWEST|FRONTIER|SPIRIT AIR"
     r"|AMERICAN AIR|VIASAT|GOGO|WYNDHAM|NYLO|MORROW HOTEL|M0RROW|RESORT|RESRT|HOTEL|INN\b|LAKEWAY|WATERSHED"
     r"|I LOVE VI|ST THOMA|AIRPORT", None, "Travel"),
    (r"SPECTRUM|COSERV|CITY OF AUSTIN|CITY OF FRISCO|ATMOS|ONCOR|TXU|RELIANT|GREEN MOUNTAIN|DIRECT ENERGY"
     r"|EARTHLINK|XFINITY|COMCAST|GOOGLE FIBER|FRONTIER COMM|WATER UTIL|ELECTRIC|UTILIT", None, "Utilities & Internet"),
    (r"CINEMARK|CINEPOLIS|AMC |ALAMO DRAFTHOUSE|STUBHUB|SEATGEEK|VIVID SEATS|TICKETMASTER|LIVE NATION|AXS"
     r"|EVENTBRITE|STATE FAIR|PICKLEBALL|APC FRISCO|TOPGOLF|DRIVING RANG|BOWL|DAVE & BUSTER|MAIN EVENT|MUSEUM"
     r"|ZOO|SIX FLAGS|CONCERT|ERYKAH|KENDRICK|PRIZEPICKS|DRAFTKINGS|FANDUEL|GAMESTOP|TRU DALLAS|CITY OF ALLEN"
     r"|T4 TACTICAL|TACTICAL|SCAT JAZZ|JAZZ|LOUNGE|FEVO|CMG ESPORTS|MLG|USA TECHNOLOGIES|ARCADE", None, "Entertainment & Events"),
    (r"CLEANERS|BARBER|HAIR|NAIL|SPA\b|MASSAGE|SALON|BEAUTY|CREPEERASE|SUPERCUTS|GREAT CLIPS|CSC SERVICEWORK"
     r"|SERVICEWORKS|LAUNDR|DAZZLINGCLEAN|CLEANING|WAXING|BLISS NAIL|ULTA|SEPHORA|BATH & BODY", None, "Personal Care"),
    (r"GYM|FITNESS|MYFITCOACH|PLANET FIT|LA FITNESS|LIFETIME|YOGA|PELOTON|LOMA VISTA|GNC|VITAMIN|BANANI", None, "Health & Fitness"),
    (r"WALGREENS|CVS|PHARM|DENTAL|DENTIST|CHIROPRACT|MYCHART|HOSP|CLINIC|MEDICAL|DOCTOR|URGENT|LABCORP|QUEST DIAG"
     r"|WARBY|OPTOM|EYE|HEALTH|MCW\d|THERAP|PEDIATR|ORTHO|DERMA", None, "Medical"),
    (r"GREAT LEARNING|PARCHMENT|UNIVERSITY|UNIV DOCS|COLLEGE|COURSERA|UDEMY|MASTERCLASS|SKILLSHARE|TUITION"
     r"|CTU |SCHOOL|LEARNING", None, "Education"),
    (r"FIVERR|UPWORK|ADROLL|ALISAVEPRO|PADDLE\.NET|PRINTFUL|IM ACADEMY|DROPSHIP|CITY TRADERS|CASH CAPITAL"
     r"|MENURAMAGIC|WEWORK|WORKBOX|HUDHOMES|LLC FILING|SECRETARY OF STATE|DHGATE|ALIEXPRESS|ALI\d", None,
     "Business & Side Hustle"),

    # ---- food --------------------------------------------------------------
    (r"STARBUCKS|DUNKIN|KOLACHE|DONUT|COFFEE|CAFE MASHGIN|JPMC|CREAMER|BAKERY|BAKE|NOTHING BUNDT|KRISPY KREME"
     r"|CRUMBL|SMOOTHIE|JUICE|WILDFLOWERCAFE|ANDY'?S|FROZEN|JENI'?S|ICE CREAM|POPS\b|POUND CAKE|WAFFLE DEN"
     r"|DUTCH BROS|SCOOTERS|PEET|BOBA|TEA\b|365 MARKET|MARKET J|UVCS|VENDING|JEEVES|VEND AT", None, "Coffee & Bakery"),
    (r"MCDONALD|JACK IN THE BOX|WHATABURGER|WENDY|CHICK-FIL-A|CHICK FIL A|SONIC|TACO BELL|BURGER KING|SUBWAY"
     r"|LITTLE CAESARS|LITTLE CA\b|DOMINO|PAPA JOHN|PIZZA HUT|MARCO'?S PIZZA|DAIRY QUEEN|BRAUMS|POPEYES|KFC"
     r"|RAISING CANE|CANES|CHURCH'?S CHICKEN|BUSH'?S CHICKEN|GOLDEN CHICK|WINGSTOP|JERSEY MIKE|JIMMY JOHN|FIREHOUSE"
     r"|WHICH WICH|IN-N-OUT|IN N OUT|FIVE GUYS|SHAKE SHACK|TACO CABANA|TACO BUENO|DEL TACO|PANDA EXPRESS|ARBY"
     r"|CHIPOTLE|QDOBA|MOE'?S|FUZZY'?S|FUZZYS|SBARRO|STEAK N SHAKE|CULVER|HOT DOG|WINGS|BURGER|CHIPS HAMBURGERS"
     r"|TORCHY|TACO|TAQUERIA|TACOS|MENOS MEXICAN|MI CASITA|EL BAJIO|PAPA JOHNS|PIZZA|CHINA WOK|CHINESE|WOK|HUNAN"
     r"|CHEF CHEN|PANDA", None, "Fast Food"),
    (r"H-E-B|\bHEB\b|H E B|KROGER|WHOLE FOODS|ALDI|RANDALLS|TOM THUMB|TRADER JOE|SPROUTS|MARKET STREET|WINCO"
     r"|FOOD LION|PUBLIX|SAFEWAY|ALBERTSONS|INSTACART|SHOP N SAVE|FIESTA|99 RANCH|H MART|CENTRAL MARKET"
     r"|MARKET MECHANICS|GROCER|SUPERMARKET|IC\* INSTACART|FOOD 4 LESS|SAM'?S CLUB|COSTCO", None, "Groceries"),
    (r"OLIVE GARDEN|WAFFLE HOUSE|IHOP|DENNY|NORMA'?S CAFE|NORMAS CAFE|CHEESECAKE|TEXAS ROADHOUSE|ON THE BORDER"
     r"|YARD HOUSE|BUFFALO WILD|APPLEBEE|CHILI'?S|RED LOBSTER|OUTBACK|LONGHORN|CRACKER BARREL|BJ'?S|TERRY BLACK"
     r"|HUTCHINS|RUDY'?S|RUDYS|BBQ|BAR-B-Q|SMOKEHOUSE|SOUL FOOD|BOBBY BS|KELLY'?S CAJUN|KELLYS CAJUN|CAJUN|SEAGER"
     r"|LEGACY HALL|MI COCINA|RAVENNA|GRILL|EATZI|LA MADELEINE|PARK BISTRO|BISTRO|ROOTS CHICKEN|TST\*|SQ \*|RESTAURANT"
     r"|KITCHEN|STEAK|SUSHI|RAMEN|PHO\b|THAI|INDIAN|ITALIAN|MEXICAN|DINER|BAR\b|PUB\b|TAVERN|BREWING|BREWERY|WINE"
     r"|CANTINA|CROWN BLOCK|CROSSING|54TH STREET|WILD DETECTIVES|FREE MAN|BLACK MEG|ASPEN|NYLO|USBFCTICHAR|DEAD FISH"
     r"|MIAMI VIBES|SIMPLY GOOD|MIGHTY FINE|LO-LOS|BAYLOR UNIV CONCESS|CONCESS|WAFFLE|EVEN\.BIZ|WILLIAMS &AMP; FUDGE"
     r"|FUDGE|AUSTIN'?S PARK|CATER", None, "Restaurants"),

    # ---- shopping ----------------------------------------------------------
    (r"WALMART|WAL-MART|WM SUPERCENTER|TARGET|FAMILY DOLLAR|DOLLAR GENERAL|DOLLAR TREE|FIVE BELOW|BIG LOTS"
     r"|MEIJER|FRED MEYER", None, "General Merchandise"),
    (r"AMAZON|AMZN|EBAY|ETSY|TEMU|SHEIN|WISH\.COM|WAYFAIR|WF \*|OVERSTOCK|SHOP APP|SHOPIFY\*|PAYPAL \*", None, "Online Shopping"),
    (r"BEST BUY|APPLE STORE|APPLE\.COM|GAMESTOP|MICRO CENTER|NEWEGG|B&H|SAMSUNG|DELL|LOGITECH", None, "Electronics & Gaming"),
    (r"HOME DEPOT|LOWE'?S|IKEA|LIVING SPACES|HOMEGOODS|BED BATH|ASHLEY|ROOMS TO GO|HARBOR FREIGHT|ACE HARDWARE"
     r"|MENARDS|CONTAINER STORE|POTTERY BARN|CRATE|WEST ELM|MICHAELS|HOBBY LOBBY|JOANN", None, "Home & Furniture"),
    (r"NORDSTROM|MACY|DILLARD|KOHL|ROSS|MARSHALLS|TJ ?MAXX|BURLINGTON|H&M|ZARA|UNIQLO|NIKE|ADIDAS|FOOT LOCKER"
     r"|POLO|RALPH LAUREN|GAP\b|OLD NAVY|FOREVER 21|NEIMAN|SAKS|ACADEMY SPORTS|DICK'?S|LULULEMON|JOURNEYS|EXPRESS\b"
     r"|TOP DAWG|KULTURAL VIBEZ|SEVENTH AVENUE|GOODWILL|SUPREME H|CLOTH|APPAREL|SHOE|BOUTIQUE|JEWEL|ARMANDO"
     r"|ESPECIALLY YOURS|ELECTEDFASH|BEAUTIFULWO", None, "Clothing & Accessories"),
    (r"BARNES|HALF PRICE BOOKS|BOOKS|KINOKUNIYA|BOOKSTORE|OFFICEMAX|OFFICE DEPOT|STAPLES|UPS STORE|USPS|FEDEX"
     r"|PARTY CITY|SPENCER|TOY|LEGO|PET|CHEWY|PETSMART|PETCO|GIFT|ENCANTO|PAWN|DALLAS MTV", None, "Books & Hobbies"),
]
RULES = [(re.compile(p), v, c) for p, v, c in R]

# Fallback: source category -> ours (for rows no rule matched)
SRC_FALLBACK = {
    "Dining & Drinks": "Restaurants", "Shopping": "General Merchandise", "Auto & Transport": "Gas & Convenience",
    "Entertainment & Rec.": "Entertainment & Events", "Bills & Utilities": "Utilities & Internet",
    "Credit Card Payment": X_CC, "Groceries": "Groceries", "Software & Tech": "Software & AI Tools",
    "Income": X_INCOME, "Internal Transfers": X_TRANSFER, "Fees": "Fees & Interest",
    "Travel & Vacation": "Travel", "Personal Care": "Personal Care", "Loan Payment": "Buy Now Pay Later",
    "Dropshipping": "Business & Side Hustle", "Tithes @ Church": "Church & Charity", "Gifts": "Payments to People",
    "Medical": "Medical", "Health & Wellness": "Health & Fitness", "Investment": X_INVEST,
    "Home & Garden": "Home & Furniture", "Cash & Checks": "Cash Withdrawals", "Education": "Education",
    "Cash Rewards": X_REWARD, "Music Production": "Music Production", "Business": "Business & Side Hustle",
    "Legal": "Insurance", "Charitable Donations": "Church & Charity", "Savings Transfer": X_TRANSFER,
    "Reimbursement": X_REWARD, "Ignore": X_TRANSFER, "Family Care": "Family Support", "Taxes": "Fees & Interest",
    "Pets": "Books & Hobbies", "Energy drink": "Gas & Convenience", "$1,000 budget": "Other", "$700 budget": "Other",
}

# Canonical vendor names for the big recurring merchants (applied after rule match)
VENDOR_CANON = [
    (r"7-ELEVEN|7ELEVEN", "7-Eleven"), (r"CIRCLE ?K", "Circle K"), (r"WALMART|WAL-MART|WM SUPERCENTER", "Walmart"),
    (r"TARGET", "Target"), (r"\bSHELL\b", "Shell"), (r"JACK IN THE BOX", "Jack in the Box"),
    (r"MCDONALD", "McDonald's"), (r"NTTA", "NTTA Tolls"), (r"UBER ?\*? ?EATS|UBEREATS", "Uber Eats"),
    (r"CHEVRON", "Chevron"), (r"TROPICAL SMOOTHIE", "Tropical Smoothie Cafe"), (r"OLIVE GARDEN", "Olive Garden"),
    (r"CEFCO", "CEFCO"), (r"WHATABURGER", "Whataburger"), (r"WAFFLE HOUSE", "Waffle House"),
    (r"SPOTIFY", "Spotify"), (r"AMAZON PRIME", "Amazon Prime"), (r"AMAZON|AMZN", "Amazon"), (r"DOMINO", "Domino's"),
    (r"EXXON", "ExxonMobil"), (r"RACETRAC|RACE TRAC", "RaceTrac"), (r"H-E-B|\bHEB\b|H E B", "H-E-B"),
    (r"TODAYS CAR WASH", "Today's Car Wash"), (r"CHICK-FIL-A|CHICK FIL A", "Chick-fil-A"), (r"WENDY", "Wendy's"),
    (r"FAMILY DOLLAR", "Family Dollar"), (r"KOLACHE HEAVEN", "Kolache Heaven"), (r"KOLACHE FACTORY", "Kolache Factory"),
    (r"UVCS", "UVCS Frisco Square"), (r"AUDIBLE", "Audible"), (r"SONIC", "Sonic"), (r"\bQT\b|QUIKTRIP", "QuikTrip"),
    (r"GODS WAY", "God's Way Church"), (r"DJ DAWN", "DJ Dawn"), (r"\bUBER\b|UBR\*", "Uber"),
    (r"JPMC|CAFE MASHGIN|AMK JPMC|CREAMER", "JPMC Office Cafe"), (r"NETFLIX HOUSE", "Netflix House"),
    (r"NETFLIX", "Netflix"), (r"REPLIT", "Replit"), (r"ALIEXPRESS|ALI\d", "AliExpress"),
    (r"YOUTUBE ?PREMIUM", "YouTube Premium"), (r"YOUTUBE|GOOGLE YOUTUBE", "YouTube"), (r"STAR MART", "Star Mart"),
    (r"365 MARKET|MARKET J", "365 Market (vending)"), (r"CREDIT PROTECT", "Credit Protection (Credit One)"),
    (r"MENOS MEXICAN", "Menos Mexican Grill"), (r"SCENTBIRD|SBD\*", "Scentbird"), (r"PATREON", "Patreon"),
    (r"BOBBY BS", "Bobby B's Soul Food"), (r"CORNER DONUTS", "Corner Donuts & Kolaches"),
    (r"LITTLE CAESARS|LITTLE CA\b", "Little Caesars"), (r"PAPA JOHN", "Papa John's"), (r"BEST BUY", "Best Buy"),
    (r"TRADEMARK CAR WASH", "Trademark Car Wash"), (r"EXPERIAN", "Experian"), (r"STARZ", "Starz"),
    (r"RANDALLS", "Randalls"), (r"TURO", "Turo"), (r"HULU|HLU\*", "Hulu"), (r"CVS", "CVS"),
    (r"O'?REILLY", "O'Reilly Auto Parts"), (r"FIVERR", "Fiverr"), (r"TRACKLIB", "Tracklib"), (r"WEWORK", "WeWork"),
    (r"U-HAUL|UHAUL", "U-Haul"), (r"MARCO'?S PIZZA", "Marco's Pizza"), (r"ATT\*|AT&T|ATT BILL|\bATT\b", "AT&T"),
    (r"ROKU", "Roku"), (r"EL BAJIO", "El Bajio"), (r"NORMA'?S CAFE|NORMAS CAFE", "Norma's Cafe"),
    (r"DALLAS MORNING NEWS", "Dallas Morning News"), (r"TAQUERIA MEXICO", "Taqueria Mexico"),
    (r"KLARNA", "Klarna"), (r"FUZZY", "Fuzzy's Taco Shop"), (r"BARNES", "Barnes & Noble"),
    (r"GOOGLE ONE", "Google One"), (r"STARBUCKS", "Starbucks"), (r"MARIANAS TACO", "Mariana's Taco Shop"),
    (r"VERIZON|VZWRLSS", "Verizon"), (r"TACO CABANA", "Taco Cabana"), (r"GOOGLE CLOUD|GOOGLE STORAGE", "Google Cloud Storage"),
    (r"TACOS Y MAS", "Tacos y Mas"), (r"MICROSOFT\*XBOX|XBOX", "Xbox"), (r"MICROSOFT|MSBILL", "Microsoft 365"),
    (r"WALGREENS", "Walgreens"), (r"KROGER", "Kroger"), (r"SHOPIFY", "Shopify"), (r"KELLY'?S CAJUN|KELLYS CAJUN", "Kelly's Cajun Grill"),
    (r"COSERV", "CoServ Electric"), (r"CLEAN GETAWAY", "Clean Getaway Car Wash"), (r"GEICO", "GEICO"),
    (r"PROGRESSIVE|PROG COUNTY", "Progressive"), (r"ROCKET MONEY|TRUEBILL", "Rocket Money"),
    (r"PARAMOUNT", "Paramount+"), (r"BEFUNKY", "BeFunky"), (r"SUBWAY", "Subway"), (r"SERATO", "Serato"),
    (r"BURGER KING", "Burger King"), (r"DOLLAR GENERAL", "Dollar General"), (r"IHOP", "IHOP"),
    (r"WHOLE FOODS", "Whole Foods"), (r"SOUNDCLOUD", "SoundCloud"), (r"CAPITAL ONE AUTO", "Capital One Auto"),
    (r"SOUNDSTRIPE", "Soundstripe"), (r"DAIRY QUEEN", "Dairy Queen"), (r"QUICK CHECK", "Quick Check"),
    (r"WSJ|WALL-ST|DOW JONES|D J\*", "Wall Street Journal"), (r"LOWE'?S", "Lowe's"), (r"JOHNS GYM", "John's Gym"),
    (r"DIAMOND SHAMROCK", "Diamond Shamrock"), (r"EATZI", "Eatzi's"), (r"SPLICE", "Splice"),
    (r"TRADINGVIEW", "TradingView"), (r"PERPLEXIT", "Perplexity"), (r"AMERICAN AIR", "American Airlines"),
    (r"FLUXTATION", "Fluxtation"), (r"SSAIYAD", "SSaiyad (beats)"), (r"GUITAR CENTER", "Guitar Center"),
    (r"ROSS DRESS", "Ross"), (r"DUNKIN", "Dunkin'"), (r"CINEMARK", "Cinemark"), (r"HOME DEPOT", "Home Depot"),
    (r"TIDAL", "TIDAL"), (r"WAYFAIR|WF \*", "Wayfair"), (r"ADOBE", "Adobe"), (r"CLAUDE|ANTHROPIC", "Claude.ai"),
    (r"NOTION", "Notion"), (r"CANVA", "Canva"), (r"SPECTRUM", "Spectrum"), (r"T-MOBILE", "T-Mobile"),
    (r"BRAUMS", "Braum's"), (r"PICKLEBALL|APC FRISCO", "Ace Pickleball Club"), (r"SEATGEEK", "SeatGeek"),
    (r"STUBHUB", "StubHub"), (r"AFFIRM", "Affirm"), (r"HILTON", "Hilton"), (r"HAMPTON INN", "Hampton Inn"),
    (r"AIRBNB", "Airbnb"), (r"ALLY (ALLY )?PAYMT|ALLY RETRY", "Ally Auto"), (r"BMW", "BMW service"), (r"GREAT LEARNING", "Great Learning"),
    (r"INTEREST CHARGE|INTEREST ON PURCHASES", "Card interest charges"), (r"LATE FEE|LATE PAYMENT", "Late fees"),
    (r"MONTHLY SERVICE FEE", "Bank monthly service fee"), (r"CASH APP\*RONDA|RONDA SEPH", "Ronda Sephus (Cash App)"),
    (r"CASH APP\*ROY|ROY SEPHUS", "Roy Sephus (Cash App)"), (r"WILDFLOWERCAFE", "Wildflower Cafe"),
    (r"CSC SERVICEWORK", "CSC ServiceWorks (laundry)"), (r"USAA INSURANCE", "USAA Insurance"), (r"NOBLR", "Noblr"),
    (r"KOHL'?S", "Kohl's"), (r"NORDSTROM", "Nordstrom"), (r"GAMESTOP", "GameStop"), (r"LYFT", "Lyft"),
    (r"TRADER JIM", "Trader Jim's Pawn"),
    (r"NFL", "NFL+"), (r"XME INC", "XME (app subscription)"), (r"CHESS", "Chess.com"), (r"MEGA LIMITED", "MEGA cloud storage"),
    (r"MAX\.COM|HBO", "Max"), (r"MYFITCOACH", "MyFitCoach"), (r"RAPIDAPI", "RapidAPI"), (r"GOOGLE WORKSPACE", "Google Workspace"),
    (r"DISNEY", "Disney+"), (r"PLANOLY", "Planoly"), (r"LUCID", "Lucid"), (r"SQUARESPACE", "Squarespace"), (r"TODOIST", "Todoist"),
    (r"PLACEIT", "Placeit"), (r"RESUME", "Resume-Now"), (r"CBS", "Paramount+ (CBS)"), (r"LOVABLE", "Lovable"), (r"VISILY", "Visily"),
    (r"APIFY", "Apify"), (r"TRANSUNION|TU \*", "TransUnion"), (r"LEXINGTON LAW", "Lexington Law"), (r"ESPN", "ESPN+"),
    (r"SOUNDCLOUD", "SoundCloud"), (r"FACEBK|FACEBOOK|META ADS", "Facebook ads"), (r"PRINTFUL", "Printful"), (r"DHGATE", "DHgate"),
]
VENDOR_CANON = [(re.compile(p), v) for p, v in VENDOR_CANON]

CITY_RE = re.compile(
    r"\b(FRISCO|DALLAS|PLANO|KILLEEN|HARKER H\w*|RICHARDSON|IRVING|ARLINGTON|AUSTIN|FORT WORTH|CARROLLTON|"
    r"LEWISVILLE|MCKINNEY|ALLEN|TEMPLE|WACO|GRAPEVINE|FARMERS BRANC\w*|THE COLONY|BELTON|SAN FRANCISCO|"
    r"LAS VEGAS|OAKLAND|TROY|HILLSBORO|PANTEGO|COPPELL|ADDISON|DENTON|GARLAND|MESQUITE|ROUND ROCK|"
    r"NEW YORK|SAN JOSE|SAN ANTONIO|HOUSTON|LITTLE ELM|PROSPER|CEDAR HIL\w*|KENNEDALE)\b.*$")


def clean_vendor(name, desc):
    s = str(name) if isinstance(name, str) and name.strip() else str(desc)
    if re.fullmatch(r"\s*\d+\s*", s):
        s = str(desc)
    s = s.upper()
    s = re.sub(r"\(CASH\)|AUTHID:\s*\d+|DEBIT CARD PURCHASE|POS DEBIT|DEBIT CRD CREDIT ADJ|DEBIT CARD REFUND"
               r"|PAYMENT RECEIPT CREDIT|ACH WITHDRAWAL|ACH DEP|WEB PMTS|- \d{4}$", " ", s)
    s = re.sub(r"^(POS\w*\s+|TST\*\s*|SQ \*|PY \*|SP \*?\s*|CKO\*|DNH\*|IC\*\s*|BM \*\s*|ZING \*\s*|CTLP\*|PAYPAL \*)", "", s)
    s = CITY_RE.sub("", s)
    s = re.sub(r"\b[A-Z0-9]*\d[A-Z0-9]{7,}\b", " ", s)          # reference codes
    s = re.sub(r"#\s*\d+|\b\d{2}/\d{2}\b|\b\d+\b|\*", " ", s)
    s = re.sub(r"\s+[A-Z]{2}$", "", s.strip())                    # trailing state
    s = re.sub(r"\s+", " ", s).strip(" -.,&")
    if not s:
        return str(name).strip()[:40]
    return s.title()[:40]


def load(path):
    df = pd.read_excel(path)
    df["Account Number"] = df["Account Number"].map(lambda v: "" if pd.isna(v) else str(v))
    n_raw = len(df)
    # 1) exact / whitespace-variant duplicates from repeated imports
    df = df.drop_duplicates(["Date", "Account Name", "Account Number", "Name", "Amount"])
    # 2) Fidelity export records outflows as negatives; flip to match every other account
    fid = df["Institution Name"].eq("Fidelity")
    df.loc[fid, "Amount"] = -df.loc[fid, "Amount"]
    # 3) the same card imported under two account numbers (NFCU 2084 / blank, Credit One 2865 / 7459)
    df["acct_group"] = df["Institution Name"] + "|" + df["Account Name"]
    key = ["acct_group", "Date", "Name", "Amount"]
    df["occ"] = df.groupby(key + ["Account Number"]).cumcount()
    df = df.sort_values("Account Number", ascending=False)  # prefer rows that carry an account number
    df = df.drop_duplicates(key + ["occ"])
    stats = {"raw_rows": n_raw, "clean_rows": len(df), "duplicates_removed": n_raw - len(df)}
    return df.sort_values("Date").reset_index(drop=True), stats


def classify(row):
    name, desc = row["Name"], row["Description"]
    u = f"{name} | {desc}".upper()
    note = str(row.get("Note") or "").upper()
    if any(k in note for k in ("BITCOIN", "ETHEREUM", "XRP")):
        return X_INVEST, clean_vendor(name, desc)
    cat = vendor = None
    for rx, v, c in RULES:
        if rx.search(u):
            cat, vendor = c, v
            break
    if cat is None:
        cat = SRC_FALLBACK.get(row["Category"], "Other")
        if cat == X_INCOME and row["Amount"] > 0:
            cat = "Other"
    # a negative amount with a spending category is a refund (it nets against the category)
    if vendor is None:
        for rx, v in VENDOR_CANON:
            if rx.search(u):
                vendor = v
                break
    if vendor is None:
        vendor = clean_vendor(name, desc)
    if cat == "Streaming & Media" and vendor == "Amazon":
        vendor = "Amazon Prime & Digital"
    return cat, vendor


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "data/transactions.xlsx"
    out = sys.argv[2] if len(sys.argv) > 2 else str(ROOT / "index.html")
    df, stats = load(src)
    res = df.apply(classify, axis=1, result_type="expand")
    df["cat"], df["vendor"] = res[0], res[1]
    df["group"] = df["cat"].map(CAT2GROUP).fillna("Excluded")
    df["spend"] = ~df["cat"].isin(EXCLUDED)

    sp = df[df["spend"]].copy()
    inc = df[df["cat"].eq(X_INCOME)].copy()

    # compact transaction list for the page (spend rows + income rows)
    cats = sorted(set(sp["cat"]))
    vendors = sorted(set(sp["vendor"]))
    accts = sorted(set(df["Account Name"].str.replace("Â", "")))
    cidx = {c: i for i, c in enumerate(cats)}
    vidx = {v: i for i, v in enumerate(vendors)}
    aidx = {a: i for i, a in enumerate(accts)}
    tx = [[d.strftime("%Y-%m-%d"), round(float(a), 2), cidx[c], vidx[v], aidx[acc.replace("Â", "")],
           str(desc)[:70]]
          for d, a, c, v, acc, desc in zip(sp["Date"], sp["Amount"], sp["cat"], sp["vendor"],
                                           sp["Account Name"], sp["Description"])]
    income = [[d.strftime("%Y-%m-%d"), round(float(-a), 2), clean_vendor(n, ds)]
              for d, a, n, ds in zip(inc["Date"], inc["Amount"], inc["Name"], inc["Description"])]

    excluded = (df[~df["spend"]].groupby("cat")["Amount"].agg(["size", "sum"]).round(2)
                .reset_index().values.tolist())
    stats.update({
        "spend_rows": int(len(sp)), "first": df["Date"].min().strftime("%Y-%m-%d"),
        "last": df["Date"].max().strftime("%Y-%m-%d"), "excluded": excluded,
        "uncategorised_share": round(float(sp.loc[sp["cat"].eq("Other"), "Amount"].sum() / sp["Amount"].sum()), 4),
    })
    data = {
        "stats": stats, "cats": cats, "vendors": vendors, "accounts": accts,
        "groups": list(GROUPS.keys()), "cat2group": CAT2GROUP, "essential": sorted(ESSENTIAL),
        "tx": tx, "income": income,
    }
    html = (ROOT / "template.html").read_text()
    payload = json.dumps(data, separators=(",", ":")).replace("</", "<\\/")
    Path(out).write_text(html.replace("/*__DATA__*/null", payload))
    print(json.dumps({k: v for k, v in stats.items() if k != "excluded"}, indent=1))
    for row in excluded:
        print("  excluded:", row)
    return df


if __name__ == "__main__":
    main()

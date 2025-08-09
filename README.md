SKU Multi-Tool | مجموعة أدوات SKU الاحترافية
<p align="center">
<img src="https://placehold.co/600x300/2E2E2E/FFFFFF?text=SKU+Multi-Tool&font=roboto" alt="Project Banner">
</p>

<p align="center">
<a href="https://github.com/Mahmoud0x2/comparer_app/releases/latest">
<img src="https://img.shields.io/github/v/release/Mahmoud0x2/comparer_app?style=for-the-badge&logo=github&color=blue" alt="Latest Release">
</a>
<img src="https://img.shields.io/github/downloads/Mahmoud0x2/comparer_app/total?style=for-the-badge&logo=github&color=green" alt="Total Downloads">
</p>


🇸🇦 العربية
نظرة عامة
مجموعة أدوات SKU الاحترافية هو تطبيق مكتبي احترافي وعالي الأداء، مصمم لتبسيط وتسريع مهام إدارة المخزون. يوفر البرنامج مجموعة من الأدوات القوية لمقارنة قوائم المنتجات وسحب البيانات، كل ذلك ضمن واجهة مستخدم عصرية وسهلة الاستخدام. هذه الأداة مثالية لمديري التجارة الإلكترونية، ومسؤولي المستودعات، ومحللي البيانات الذين يتعاملون مع ملفات إكسل كبيرة.

✨ الميزات الرئيسية
أداتان قويتان في برنامج واحد:

مقارنة SKUs: تقوم بمقارنة ذكية بين قائمتين من رموز SKU (من نفس الملف أو ملفين مختلفين) وتقدم ملخصًا تفصيليًا للتطابقات المؤكدة، والتقريبية، والعناصر غير المتطابقة.

سحب بيانات المنتجات: تبحث بسرعة عن قائمة منتجات داخل صفحة بيانات رئيسية وتقوم بسحب جميع الصفوف المطابقة ببياناتها الكاملة.

أداء فائق السرعة:

يستخدم تقنية تعدد الخيوط (Multi-threading) لتشغيل جميع عمليات المعالجة الثقيلة في الخلفية، مما يحافظ على استجابة الواجهة في جميع الأوقات.

يعتمد على خوارزميات محسّنة، بما في ذلك تقنية الفهرسة الذكية للمطابقة التقريبية، للتعامل مع مجموعات البيانات الكبيرة بكفاءة.

واجهة مستخدم عصرية وسهلة:

واجهة أنيقة ذات طابع داكن مدعومة بمكتبة sv-ttk، مستوحاة من أنظمة التشغيل الحديثة.

بطاقات إحصائيات فورية: احصل على نظرة عامة فورية على النتائج من خلال بطاقات ملخصة واضحة وموجزة.

مؤشر تقدم في الخلفية: شريط تقدم غير متطفل يوضح أن التطبيق يعمل دون إبطاء العملية.

مرونة وسهولة في الاستخدام:

اختصار "نفس الملف الأول": قارن بسهولة بين صفحتين أو عمودين مختلفين داخل نفس الملف دون الحاجة إلى اختياره مرتين.

خيارات حفظ مرنة: احفظ التقرير النهائي كملف Excel (.xlsx) أو CSV (.csv) جديد، أو أضفه كصفحة جديدة إلى أحد الملفات الأصلية.

🚀 كيفية الاستخدام
تحميل البرنامج:

اذهب إلى صفحة الإصدارات (Releases) الموجودة على يمين الصفحة الرئيسية للمشروع.

قم بتحميل ملف app.exe من أحدث إصدار.

تشغيل البرنامج:

لا حاجة للتثبيت! فقط انقر نقرًا مزدوجًا على ملف app.exe الذي تم تحميله لتشغيله.

اختر الأداة:

اختر بين "مقارنة SKUs" و "سحب بيانات المنتجات" باستخدام التبويبات في الأعلى.

اضبط الإعدادات وشغّل:

اتبع التعليمات التي تظهر على الشاشة لاختيار الملفات والصفحات والأعمدة.

اضغط على زر "بدء" ودع الأداة تقوم بالعمل الشاق.

عند الانتهاء، سيطلب منك البرنامج حفظ النتائج.

🛠️ التقنيات المستخدمة
Python 3.10

Tkinter مع sv-ttk للواجهة العصرية.

Pandas لمعالجة البيانات عالية الأداء.

thefuzz للمطابقة التقريبية للسلاسل النصية.

Nuitka لتحويل الكود إلى ملف تنفيذي مستقل.

GitHub Actions لعمليات البناء والإصدار الآلية.

🇬🇧 English
Overview
SKU Multi-Tool is a professional, high-performance desktop application designed to simplify and accelerate inventory management tasks. It provides a suite of powerful tools for comparing product lists and fetching data, all wrapped in a modern and intuitive user interface. This tool is perfect for e-commerce managers, warehouse operators, and data analysts who work with large Excel files.

✨ Key Features
Two Powerful Tools in One:

SKU Comparer: Intelligently compares two lists of SKUs (from the same or different files) and provides a detailed summary of confirmed matches, fuzzy matches, and unmatched items.

Product Data Fetcher: Quickly searches for a list of items within a master data sheet and extracts all corresponding rows with their complete data.

Blazing-Fast Performance:

Utilizes multi-threading to run all heavy processing in the background, keeping the UI responsive at all times.

Employs optimized algorithms, including an intelligent indexing technique for fuzzy matching, to handle large datasets efficiently.

Modern & Intuitive UI:

A sleek, dark-themed interface powered by the sv-ttk library, inspired by modern operating systems.

Instant Stat Cards: Get an immediate overview of the results through clear and concise summary cards.

Background Progress Indicator: A non-intrusive progress bar shows that the application is working without slowing down the process.

Flexible & Convenient:

"Use Same File" Shortcut: Easily compare two different sheets or columns within the same file without selecting it twice.

Flexible Save Options: Save the final report as a new Excel (.xlsx) or CSV (.csv) file, or add it as a new sheet to one of the original files.

🚀 How to Use
Download the Application:

Go to the Releases Page on the right side of the project's main page.

Download the app.exe file from the latest release.

Run the Application:

No installation is needed! Simply double-click the downloaded app.exe file to run it.

Select a Tool:

Choose between "SKU Comparer" and "Product Data Fetcher" using the tabs at the top.

Configure and Run:

Follow the on-screen instructions to select your files, sheets, and columns.

Click the "Start" button and let the tool do the heavy lifting.

Once finished, you will be prompted to save the results.

🛠️ Built With
Python 3.10

Tkinter with sv-ttk for the modern UI.

Pandas for high-performance data manipulation.

thefuzz for fuzzy string matching.

Nuitka for compiling into a standalone executable.

GitHub Actions for automated builds and releases.

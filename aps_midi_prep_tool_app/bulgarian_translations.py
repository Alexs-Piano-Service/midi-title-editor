"""Bulgarian translations for the final message and common-text catalogs.

Technical product names, format names, placeholders, and HTML structure are
kept intact so these translations can be applied as late catalog overrides.
"""

from __future__ import annotations


BULGARIAN_MESSAGE_TRANSLATIONS = {
    "settings.reset_hidden_dialogs.done.title": "Скритите диалогови прозорци са нулирани",
    "settings.reset_hidden_dialogs.done.message": (
        "Предупрежденията, потвържденията, напомнянията за обновления и отчетите "
        "с карти на секторите от Greaseweazle отново ще се показват."
    ),
    "dialog.confirm_image_save.title": "Потвърждение за запис на образ",
    "dialog.confirm_floppy_save.title": "Потвърждение за запис на дискета",
    "dialog.confirm_save_update.message": (
        'При запис ще бъдат изтрити окончателно {count} {entry_word} от {container}.\n\nПреименуваните файлове ще запазят данните на песните под новото име.\n\nДа се продължи ли?'
    ),
    "dialog.save_to_floppy.hint": (
        'Изберете устройството с форматирана дискета за текущите файлове във формат {format}.'
    ),
    "dialog.write_image_to_floppy.hint": (
        "Изберете къде да бъде записан текущият образ във формат {format}."
    ),
    "dialog.about.title": "Относно {app}",
    "dialog.disclaimer.html": (
        "<p>APS MIDI Prep Tool се предоставя за законни дейности по съхраняване, "
        "поправка и осигуряване на съвместимост.</p>"
        "<p>Когато е възможно, работете с копия, пазете резервни копия и "
        "проверявайте резултатите, преди да разчитате на тях. Вие носите "
        "отговорност за загуба на данни, повреда на диск, поведение на инструмент "
        "или други последствия от използването на софтуера.</p>"
        "<p>Използвайте инструмента само с дискове и файлове, които притежавате "
        "или за които имате разрешение да съхранявате, преобразувате, поправяте "
        "или променяте. Не го използвайте за разпространение на защитена с "
        "авторски права музика, търговски библиотеки за механично пиано, "
        "собственически софтуер или други материали, които нямате право да "
        "споделяте.</p>"
        "<p>Този софтуер е независим и не е свързан, спонсориран или одобрен от "
        "Yamaha, Disklavier, PianoSoft, Electone, Clavinova, PianoDisc, Nalbantov, "
        "Greaseweazle, Akai, MPC или други споменати компании и продукти. Имената "
        "на трети страни се използват само за обозначаване на съвместими формати, "
        "носители, инструменти и работни процеси. Търговските марки и имената на "
        "продукти принадлежат на съответните им собственици.</p>"
        "<p><strong>Правила на Alex's Piano Service LLC:</strong><br>"
        "<a href=\"https://www.alexanderpeppe.com/disclaimer/\">Отказ от "
        "отговорност</a> &nbsp;|&nbsp; "
        "<a href=\"https://www.alexanderpeppe.com/privacy-policy/\">Политика за "
        "поверителност</a> &nbsp;|&nbsp; "
        "<a href=\"https://www.alexanderpeppe.com/dmca-policy/\">Политика по "
        "DMCA</a></p>"
    ),
    "gw.sector.no_map": (
        "За тази операция няма налична карта на секторите от Greaseweazle."
    ),
    "gw.sector.legend.ok": "Изправен",
    "guidance.command_missing": (
        "Инсталирайте липсващата команда или използвайте комплектована версия, "
        "след което опитайте отново."
    ),
    "guidance.windows_device_access": (
        "В Windows затворете другите дискови инструменти и стартирайте приложението "
        "с администраторски права, ако е необходим пряк достъп до устройството."
    ),
    "guidance.cancelled": (
        'Операцията беше отменена преди завършването ѝ.'
    ),
    "guidance.greaseweazle_sector_failures": (
        "Проверете устройството, кабела, формата на диска и състоянието на носителя. "
        "Опитайте с повече повторения или с архивен SCP запис, ако дискът може да е "
        "повреден или защитен от копиране."
    ),
    "guidance.write_protected": (
        "Проверете плъзгача за защита от запис на диска и настройките за защита от "
        "запис в приложението, преди да опитате отново."
    ),
}


BULGARIAN_TEXT_TRANSLATIONS = {
    "Convert": "Преобразуване",
    "SMF1 -> SMF0": "SMF1 → SMF0",
    "E-SEQ -> MIDI": "E-SEQ → MIDI",
    "MIDI -> E-SEQ": "MIDI → E-SEQ",
    "For Save As folder exports, create a subfolder from the catalog number and album title. Save As Image and floppy writes are not affected.": (
        "При експортиране в папка чрез „Запис като“ създава подпапка от "
        "каталожния номер и заглавието на албума. „Запис като образ“ и записването "
        "на дискета не се засягат."
    ),
    "Create an album subfolder only for Save As folder exports.": (
        "Създава подпапка за албума само при експортиране в папка чрез „Запис като“."
    ),
    "Please wait for the current operation to finish before changing album subfolder output.": (
        "Изчакайте текущата операция да завърши, преди да промените извеждането в "
        "подпапка на албума."
    ),
    "Saved in album subfolder: {folder}": "Записано в подпапката на албума: {folder}",
    "Create Album Subfolder is on, but no album title or catalog number is available, so files were saved directly in the selected folder.": (
        "Създаването на подпапка за албума е включено, но няма заглавие на албум "
        "или каталожен номер, затова файловете са записани направо в избраната папка."
    ),
    "Files have been saved to the new folder.": "Файловете са записани в новата папка.",
    "Saving files to new folder...": "Записване на файловете в новата папка...",
    "Preparing exported files...": "Подготовка на файловете за експортиране...",
    "Metadata summary written to {filename}.": (
        "Обобщението на метаданните е записано във {filename}."
    ),
    ".tags.txt sidecar file(s) were written next to the exported files.": (
        "Помощните .tags.txt файлове са записани до експортираните файлове."
    ),
    "ImagePath": 'Път в образа',
    "New Image": "Нов образ",
    "Save To Floppy": "Запис на дискета",
    "Auto Write-Protect": "Автоматична защита от запис",
    "Write-Protect Original": "Защита на оригинала от запис",
    "Create Tag Sidecars When Saving": "Създаване на помощни файлове с етикети при запис",
    "Create Metadata Summary When Saving": "Създаване на обобщение на метаданните при запис",
    "Hide the Album Info panel, including Album Title, Catalog Number, and Create Album Subfolder.": (
        "Скрива панела „Информация за албума“, включително „Заглавие на албума“, "
        "„Каталожен номер“ и „Създаване на подпапка за албума“."
    ),
    "Send Feedback": "Изпращане на обратна връзка",
    "Send Feedback...": "Изпращане на обратна връзка...",
    "Send feedback with app details and optional recent console logs.": (
        "Изпраща обратна връзка с данни за приложението и по желание скорошни "
        "регистрационни записи от конзолата."
    ),
    "Tell us what would make APS MIDI Prep Tool better, or what is working well.": (
        "Споделете какво би подобрило APS MIDI Prep Tool или какво работи добре."
    ),
    "Feedback includes app details. Logs are optional and may include recent console output and file paths.": (
        "Обратната връзка включва данни за приложението. Регистрационните записи са "
        "по желание и може да съдържат скорошен изход от конзолата и пътища до файлове."
    ),
    "What would you like to share?": "Какво бихте искали да споделите?",
    "Adds recent console output if it helps explain your feedback.": (
        "Добавя скорошен изход от конзолата, ако той помага да обясните обратната си връзка."
    ),
    "Feedback Needs Detail": "Обратната връзка се нуждае от подробности",
    "Add a short summary or feedback details before sending.": (
        "Добавете кратко обобщение или подробности към обратната връзка, преди да я изпратите."
    ),
    "No feedback endpoint is configured for this build.": (
        "За тази версия не е настроен адрес за получаване на обратна връзка."
    ),
    "Please wait for the current feedback to finish sending.": (
        "Изчакайте изпращането на текущата обратна връзка да завърши."
    ),
    "Sending feedback...": "Изпращане на обратна връзка...",
    "Feedback Sent": "Обратната връзка е изпратена",
    "Feedback sent.": "Обратната връзка е изпратена.",
    "Feedback Failed": "Неуспешно изпращане на обратната връзка",
    "Feedback failed. See View > View Logs for details.": (
        "Обратната връзка не беше изпратена. За подробности вижте „Изглед > Преглед на логове“."
    ),
    "The app could not send feedback": "Приложението не можа да изпрати обратната връзка",
    "Error message shown by the app:": "Съобщение за грешка, показано от приложението:",
    "Send a bug report with app details and optional recent console logs.": (
        "Изпраща сигнал за грешка с данни за приложението и по желание скорошни "
        "регистрационни записи от конзолата."
    ),
    "Tell us what happened and what you expected instead.": (
        "Опишете какво се случи и какво очаквахте да се случи."
    ),
    "The report includes app details. Logs are optional and may include recent console output and file paths.": (
        "Сигналът включва данни за приложението. Регистрационните записи са по "
        "желание и може да съдържат скорошен изход от конзолата и пътища до файлове."
    ),
    "Send a bug report to APS MIDI Prep Tool support. Include what you were doing and what went wrong.": (
        "Изпратете сигнал за грешка до поддръжката на APS MIDI Prep Tool. Опишете "
        "какво правехте и какво се обърка."
    ),
    "Reports may include app version, operating system details, current mode, file names or paths, and recent console output.": (
        "Сигналите може да включват версията на приложението, данни за операционната "
        "система, текущия режим, имена или пътища на файлове и скорошен изход от конзолата."
    ),
    "Short summary": "Кратко обобщение",
    "What happened? What did you expect instead?": "Какво се случи? Какво очаквахте вместо това?",
    "Optional email or contact info": "Незадължителен имейл или данни за контакт",
    "Summary": "Обобщение",
    "Contact": "Контакт",
    "When checked, the report includes the most recent console output.": (
        "Когато е отметнато, сигналът включва най-новия изход от конзолата."
    ),
    "Adds recent console output to help diagnose the problem.": (
        "Добавя скорошен изход от конзолата, за да помогне при диагностицирането на проблема."
    ),
    "Sending bug report": "Изпращане на сигнал за грешка",
    "Sending Bug Report": "Изпращане на сигнал за грешка",
    "Bug Report Not Configured": "Сигналите за грешки не са настроени",
    "No bug report endpoint is configured for this build.": (
        "За тази версия не е настроен адрес за получаване на сигнали за грешки."
    ),
    "Bug Report In Progress": "Изпраща се сигнал за грешка",
    "Please wait for the current bug report to finish sending.": (
        "Изчакайте изпращането на текущия сигнал за грешка да завърши."
    ),
    "Can't connect to Alex's Piano Service. Please check your internet connection.": (
        "Неуспешно свързване с Alex's Piano Service. Проверете връзката си с интернет."
    ),
    "Check your internet connection, then try again. You can also use View > View Logs... to save the log manually.": (
        "Проверете връзката си с интернет и опитайте отново. Можете също да използвате "
        "„Изглед > Преглед на логове...“, за да запишете журнала ръчно."
    ),
    "Original floppy write is protected. Turn off File > Write Protection > Write-Protect Original, or use Save As or Save As Image.": (
        "Записът върху оригиналната дискета е защитен. Изключете „Файл > Защита от "
        "запис > Защита на оригинала от запис“ или използвайте „Запис като“ или "
        "„Запис като образ“."
    ),
    "Original image write is protected. Turn off File > Write Protection > Write-Protect Original, or use Save As or Save As Image.": (
        "Записът върху оригиналния образ е защитен. Изключете „Файл > Защита от "
        "запис > Защита на оригинала от запис“ или използвайте „Запис като“ или "
        "„Запис като образ“."
    ),
    "Save To Image Is Off": "Записът в образ е изключен",
    "Save To Floppy Is Off": "Записът на дискета е изключен",
    "Use Save As to export files, use Save As Image to create a separate image, or turn off File > Write Protection > Write-Protect Original.": (
        "Използвайте „Запис като“, за да експортирате файловете, „Запис като образ“, "
        "за да създадете отделен образ, или изключете „Файл > Защита от запис > "
        "Защита на оригинала от запис“."
    ),
    "Use Save As Image to save an image file, or turn off File > Write Protection > Write-Protect Original.": (
        "Използвайте „Запис като образ“, за да запишете файл с образ, или изключете "
        "„Файл > Защита от запис > Защита на оригинала от запис“."
    ),
    "Song List": "Списък с песни",
    "Recover Damaged Image": "Възстановяване на повреден образ",
    "Rename All to DOS 8.3": "Преименуване на всички по DOS 8.3",
    "Trim leading/trailing title spaces and collapse repeated spaces for every listed title.": (
        "Премахва интервалите в началото и края и свежда повтарящите се интервали "
        "до един във всички заглавия от списъка."
    ),
    "Please wait for the current operation to finish before trimming title spaces.": (
        "Изчакайте текущата операция да завърши, преди да почистите интервалите в заглавията."
    ),
    "No listed titles need spacing cleanup.": (
        "Нито едно заглавие от списъка не се нуждае от почистване на интервалите."
    ),
    "No listed MIDI or E-SEQ titles need spacing cleanup.": (
        "Нито едно MIDI или E-SEQ заглавие от списъка не се нуждае от почистване на интервалите."
    ),
    "Trim Titles Not Needed": "Не е необходимо почистване на заглавията",
    "No listed titles needed spacing cleanup.": (
        "Нито едно заглавие от списъка не се нуждаеше от почистване на интервалите."
    ),
    "Trim Titles Failed": "Неуспешно почистване на заглавията",
    "Some titles could not be cleaned up": "Някои заглавия не можаха да бъдат почистени",
    "Nothing has been written yet; review the listed files and try again": (
        "Все още нищо не е записано; прегледайте файловете в списъка и опитайте отново"
    ),
    "Convert All SMF1 to SMF0": "Преобразуване на всички от SMF1 в SMF0",
    "Convert All E-SEQ to MIDI": "Преобразуване на всички от E-SEQ в MIDI",
    "Convert All MIDI to E-SEQ": "Преобразуване на всички от MIDI в E-SEQ",
    "Format Floppy Disk": "Форматиране на дискета",
    "Check for Updates": "Проверка за обновления",
    "OK": "ОК",
    "Browse": "Избор",
    "Copy to Clipboard": "Копиране в клипборда",
    "Convert to MIDI": "Преобразуване в MIDI",
    "Convert and Exit": "Преобразуване и изход",
    "Not Now": "Не сега",
    "Read": "Прочитане",
    "Recover": "Възстановяване",
    "Image": 'Създаване на образ',
    "Read using": "Четене чрез",
    "Floppy Drive": "Флопи устройство",
    "Floppy drive": "Флопи устройство",
    "Greaseweazle device": "Устройство Greaseweazle",
    "Drive": "Устройство",
    "Disk format": "Формат на диска",
    "Disk size": "Размер на диска",
    "Recovery disk format": "Формат на диска за възстановяване",
    "Read revs": "Обороти при четене",
    "Read retries": "Повторни опити при четене",
    "Save image": "Запис на образ",
    "Image type": "Тип на образа",
    "Format using": "Форматиране чрез",
    "Image using": "Създаване на образ чрез",
    "Write using": "Запис чрез",
    "No supported floppy drive detected": "Не е открито поддържано флопи устройство",
    "No Greaseweazle device detected": "Не е открито устройство Greaseweazle",
    "Start in recovery mode": "Стартиране в режим на възстановяване",
    "Normal read uses fast file-level reading when possible.": (
        "При възможност нормалното четене използва бърз достъп на ниво файлова система."
    ),
    "Recovery may take a long time and opens recovered data in a new editable image copy.": (
        "Възстановяването може да отнеме много време и отваря възстановените данни "
        "в ново редактируемо копие на образа."
    ),
    "Recovery copies the selected full disk size first; most Yamaha Disklavier floppies are IBM 720K DD.": (
        "Възстановяването първо копира целия избран размер на диска; повечето дискети "
        "за Yamaha Disklavier са IBM 720K DD."
    ),
    "Create E-SEQ disk with empty PIANODIR.FIL": "Създаване на E-SEQ диск с празен PIANODIR.FIL",
    "Do not offer to save an image": "Да не се предлага запис на образ",
    "IMG raw sector image": "IMG образ със сурови сектори",
    "BIN raw sector image": "BIN образ със сурови сектори",
    "IMA raw sector image": "IMA образ със сурови сектори",
    "Refresh": "Обновяване",
    "Format": "Форматиране",
    "More info...": "Още информация...",
    "Volume label": "Етикет на тома",
    "Current Contents": "Текущо съдържание",
    "Item": "Елемент",
    "Recommended: use a double-density disk and format it as IBM 720K DD for Yamaha Disklavier compatibility.": (
        "Препоръка: използвайте диск с двойна плътност и го форматирайте като IBM "
        "720K DD за съвместимост с Yamaha Disklavier."
    ),
    "Create an image file from a physical floppy without opening, scanning, repairing, or converting its contents.": (
        "Създава файл с образ от физическа дискета, без да отваря, сканира, поправя "
        "или преобразува съдържанието ѝ."
    ),
    "Choose SCP for a raw flux capture. HFE is the usual Nalbantov-friendly image type.": (
        "Изберете SCP за запис на суров магнитен поток. HFE е обичайният тип образ, "
        "подходящ за Nalbantov."
    ),
    "not mounted": "не е монтирано",
    "Current mount points": "Текущи точки на монтиране",
    "This device is read-only and cannot be formatted.": (
        "Това устройство е само за четене и не може да бъде форматирано."
    ),
    "Used": "Използвано",
    "Free": "Свободно",
    "Unallocated": "Неразпределено",
    "Unknown": "Неизвестно",
    "No readable usage": 'Няма данни за използваното място',
    "No readable volumes or partitions were detected.": (
        "Не са открити четими томове или дялове."
    ),
    "Label": "Етикет",
    "Mounted": "Монтирано",
    "No mounted file-system details": "Няма данни за монтирана файлова система",
    "No top-level files could be shown": "Не могат да бъдат показани файлове от най-горното ниво",
    "Unmounted or unreadable": "Демонтирано или нечетимо",
    "Device": "Устройство",
    "Use": "Използване",
    "This cannot be undone. The existing partition table, files, and disk contents will be removed.": (
        "Това действие не може да бъде отменено. Съществуващата таблица на дяловете, "
        "файловете и съдържанието на диска ще бъдат премахнати."
    ),
    "About": "Относно",
    "Author": "Автор",
    "Project website.": "Уебсайт на проекта.",
    "Create a fresh editable floppy image.": "Създава нов редактируем образ на дискета.",
    "List all image types": "Показване на всички типове образи",
    "E-SEQ disk with empty PIANODIR.FIL": "E-SEQ диск с празен PIANODIR.FIL",
    "Adds an empty Yamaha PIANODIR.FIL so the formatted disk opens in E-SEQ mode.": (
        "Добавя празен Yamaha PIANODIR.FIL, за да се отваря форматираният диск в режим E-SEQ."
    ),
    "CLI default": "По подразбиране за CLI",
    "Number of revolutions to read per track. Use 0 for Greaseweazle's default.": (
        "Брой обороти за прочитане на всяка писта. Използвайте 0 за стойността по "
        "подразбиране на Greaseweazle."
    ),
    "Number of retries per seek-retry. Use 0 for Greaseweazle's default.": (
        'Повторни опити за четене след всяко преместване на главата. Използвайте 0 за стойността по подразбиране на Greaseweazle.'
    ),
    "Choose the image type to offer after the disk opens. SCP reads as raw flux first; other types use the selected disk format.": (
        "Изберете типа образ, който да бъде предложен след отварянето на диска. SCP "
        "първо се прочита като суров магнитен поток; другите типове използват избрания "
        "формат на диска."
    ),
    "Choose SCP for a raw flux capture. Other image types are decoded using the selected disk format.": (
        "Изберете SCP за запис на суров магнитен поток. Другите типове образи се "
        "декодират чрез избрания формат на диска."
    ),
    "Copies a full disk image and tries Yamaha/FAT repair plus raw MIDI/E-SEQ/PIANODIR scanning. The source floppy is not modified.": (
        "Копира пълен образ на диска и опитва поправка на Yamaha/FAT структурата, "
        "както и сканиране на сурови MIDI/E-SEQ/PIANODIR данни. Оригиналната дискета "
        "не се променя."
    ),
    "After the floppy opens, queue detected Yamaha E-SEQ songs for Standard MIDI conversion.": (
        "След отварянето на дискетата добавя откритите Yamaha E-SEQ песни в опашката "
        "за преобразуване в Standard MIDI."
    ),
    "Choose a floppy image": "Избор на образ на дискета",
    "Choose a floppy image to recover.": "Изберете образ на дискета за възстановяване.",
    "Autodetect": "Автоматично разпознаване",
    "Image:": "Образ:",
    "Image format": "Формат на образа",
    "image format": "формат на образа",
    "Choose Damaged Floppy Image": "Избор на повреден образ на дискета",
    "Floppy Recovery Disk Size": "Размер на диска за възстановяване на дискета",
    "Choose the format of the disk in the drive. Most Yamaha Disklavier floppies are IBM 720K DD; recovery will copy exactly the selected amount of data.": (
        "Изберете формата на диска в устройството. Повечето дискети за Yamaha "
        "Disklavier са IBM 720K DD; възстановяването ще копира точно избраното "
        "количество данни."
    ),
    "of": "от",
    "SoundFont": "SoundFont банка",
    "SoundFont:": "SoundFont банка:",
}

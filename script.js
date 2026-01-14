// ================== متغيرات عامة ==================
let excelData = null;
let templateFile = null;
let excelFileName = '';
let templateFileName = 'Att NoteClinique.docx';


// ================== تحميل قالب Word تلقائياً ==================
async function loadTemplateAutomatically() {
    try {
        const response = await fetch('Att NoteClinique.docx');

        if (!response.ok) {
            throw new Error('ملف القالب غير موجود');
        }

        templateFile = await response.arrayBuffer();

        document.getElementById('template-filename').textContent = templateFileName;
        document.getElementById('template-info').style.display = 'flex';
        document.getElementById('template-label').innerHTML =
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor">' +
            '<polyline points="20 6 9 17 4 12"></polyline></svg> القالب جاهز تلقائياً';

        showAlert('تم تحميل قالب Word تلقائياً', 'success');

    } catch (error) {
        console.error(error);
        showAlert('خطأ: تأكد أن ملف Att NoteClinique.docx موجود في نفس المجلد', 'error');
    }
}


// ================== رفع ملف Excel ==================
document.getElementById('excel_file').addEventListener('change', async function (e) {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.match(/\.(xlsx|xls)$/)) {
        showAlert('يرجى اختيار ملف Excel بصيغة .xlsx أو .xls', 'error');
        e.target.value = '';
        return;
    }

    showLoading(true);

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        if (jsonData.length === 0) {
            throw new Error('ملف Excel فارغ');
        }

        const requiredColumns = ['CNE', 'NOM', 'Prénom'];
        const columns = Object.keys(jsonData[0]);
        const missing = requiredColumns.filter(c => !columns.includes(c));

        if (missing.length > 0) {
            throw new Error('أعمدة ناقصة: ' + missing.join(', '));
        }

        excelData = jsonData;
        excelFileName = file.name;

        document.getElementById('filename').textContent = file.name;
        document.getElementById('rows-count').textContent = jsonData.length;
        document.getElementById('file-info').style.display = 'flex';
        document.getElementById('file-label').innerHTML =
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor">' +
            '<polyline points="20 6 9 17 4 12"></polyline></svg> تم الرفع بنجاح';

        showAlert('تم تحميل ملف Excel بنجاح', 'success');

    } catch (error) {
        showAlert('خطأ في قراءة Excel: ' + error.message, 'error');
        e.target.value = '';
    } finally {
        showLoading(false);
    }
});


// ================== إنشاء المستند ==================
async function generateDocument() {

    if (!templateFile) {
        showAlert('ملف القالب غير محمّل', 'error');
        return;
    }

    if (!excelData) {
        showAlert('يرجى رفع ملف Excel أولاً', 'error');
        return;
    }

    const cneInput = document.getElementById('cne_input');
    const cne = cneInput.value.trim();

    if (!cne) {
        showAlert('يرجى إدخال رقم CNE', 'error');
        cneInput.focus();
        return;
    }

    showLoading(true);
    hideAlert();
    document.getElementById('result-section').style.display = 'none';

    try {
        const student = excelData.find(row =>
            String(row.CNE).trim().toLowerCase() === cne.toLowerCase()
        );

        if (!student) {
            throw new Error('لم يتم العثور على الطالب');
        }

        const today = new Date();
        const dateStr =
            String(today.getDate()).padStart(2, '0') + '/' +
            String(today.getMonth() + 1).padStart(2, '0') + '/' +
            today.getFullYear();

        const context = {
            date: dateStr,
            nom: String(student.NOM || '').trim(),
            Prénom: String(student['Prénom'] || '').trim(),
            CNE: String(student.CNE || '').trim(),

            Chirurgie: String(student.Chirurgie_Epreuve || ''),
            CHir_1Session: String(student.Chirurgie_Session || ''),
            CHir_1Note_ecrit: String(student['Chirurgie_Note ecrit'] || ''),
            CHir_1Note_malade: String(student['Chirurgie_Note malade'] || ''),

            Gynécologie: String(student['Gynécologie_Epreuve'] || ''),
            Gynéco_1Session: String(student['Gynécologie_Session'] || ''),
            Gynéco_1Note_ecrit: String(student['Gynécologie_Note ecrit'] || ''),
            Gynéco_1Note_malade: String(student['Gynécologie_Note malade'] || ''),

            Médecine: String(student['Médecine_Epreuve'] || ''),
            MED_1Session: String(student['Médecine_Session'] || ''),
            MED_1Note_ecrit: String(student['Médecine_Note ecrit'] || ''),
            MED_1Note_malade: String(student['Médecine_Note malade'] || ''),

            Pédiatrie: String(student['Pédiatrie_Epreuve'] || ''),
            PED_1Session: String(student['Pédiatrie_Session'] || ''),
            PED_1Note_ecrit: String(student['Pédiatrie_Note ecrit'] || ''),
            PED_1Note_malade: String(student['Pédiatrie_Note malade'] || '')
        };

        const zip = new PizZip(templateFile);
        const doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true
        });

        doc.render(context);

        const blob = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        const fileName = `Attestation_${context.nom}_${context.Prénom}_${context.CNE}.docx`;
        saveAs(blob, fileName);

        document.getElementById('student-name').textContent =
            `الطالب: ${context.nom} ${context.Prénom}`;

        document.getElementById('result-section').style.display = 'block';

        showAlert('تم إنشاء الشهادة بنجاح', 'success');

    } catch (error) {
        console.error(error);
        showAlert('خطأ: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}


// ================== أدوات مساعدة ==================
function showLoading(show) {
    document.getElementById('loading').style.display = show ? 'block' : 'none';
    document.getElementById('generate-btn').disabled = show;
}

function showAlert(message, type) {
    const box = document.getElementById('alert-box');
    box.textContent = message;
    box.className = `alert alert-${type}`;
    box.style.display = 'block';

    setTimeout(hideAlert, 7000);
}

function hideAlert() {
    document.getElementById('alert-box').style.display = 'none';
}


// ================== Enter ==================
document.getElementById('cne_input').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        generateDocument();
    }
});


// ================== تحميل تلقائي عند فتح الصفحة ==================
window.addEventListener('load', function () {
    loadTemplateAutomatically();
    showAlert('ارفع ملف Excel ثم أدخل CNE', 'warning');
});

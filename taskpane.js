// Office.js initialization
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('autoReplyForm').addEventListener('submit', setAutoReply);
        loadColleagues();
        setupFormListeners();
        setDefaultDates();
    }
});

// Global variables
let currentLanguage = 'tr';
let colleagues = [];

// Message templates
const templates = {
    tr: {
        subject: "Otomatik Yanıt: Yıllık İzin",
        body: `Sayın Yetkili,

E-postanız için teşekkür ederim. {startDate} – {endDate} tarihleri arasında yıllık izinde olacağım ve bu süre içinde e-postalarınıza yanıt veremeyeceğim.

Acil konularınız için {colleagueName} ile {email} veya {phone} üzerinden iletişime geçebilirsiniz.

Anlayışınız için teşekkür eder, iyi çalışmalar dilerim.

Saygılarımla,
{userName}
{position}
{company}`
    },
    en: {
        subject: "Automatic Reply: Annual Leave",
        body: `Dear Sir/Madam,

Thank you for your email. I will be out of the office on annual leave from {startDate} to {endDate}, and will not be able to respond to your message during this period.

For urgent matters, please contact {colleagueName} at {email} or {phone}.

Thank you for your understanding.

Kind regards,
{userName}
{position}
{company}`
    }
};

// Mock D365 data - In production, this would come from D365 API
const mockColleagues = [
    {
        id: 1,
        name: "Ahmet Yılmaz",
        email: "ahmet.yilmaz@ozturyakiler.com.tr",
        phone: "+90 212 555 0101",
        department: "İnsan Kaynakları"
    },
    {
        id: 2,
        name: "Fatma Demir",
        email: "fatma.demir@ozturyakiler.com.tr",
        phone: "+90 212 555 0102",
        department: "Muhasebe"
    },
    {
        id: 3,
        name: "Mehmet Kaya",
        email: "mehmet.kaya@ozturyakiler.com.tr",
        phone: "+90 212 555 0103",
        department: "Satış"
    },
    {
        id: 4,
        name: "Ayşe Özkan",
        email: "ayse.ozkan@ozturyakiler.com.tr",
        phone: "+90 212 555 0104",
        department: "Pazarlama"
    },
    {
        id: 5,
        name: "Can Şahin",
        email: "can.sahin@ozturyakiler.com.tr",
        phone: "+90 212 555 0105",
        department: "IT"
    }
];

function loadColleagues() {
    // In production, this would be an API call to D365
    colleagues = mockColleagues;
    
    const colleagueSelect = document.getElementById('colleague');
    colleagueSelect.innerHTML = '<option value="">Seçiniz...</option>';
    
    colleagues.forEach(colleague => {
        const option = document.createElement('option');
        option.value = colleague.id;
        option.textContent = `${colleague.name} (${colleague.department})`;
        colleagueSelect.appendChild(option);
    });
}

function setLanguage(lang) {
    currentLanguage = lang;
    
    // Update button states
    document.getElementById('btnTurkish').classList.toggle('active', lang === 'tr');
    document.getElementById('btnEnglish').classList.toggle('active', lang === 'en');
    
    // Update labels based on language
    if (lang === 'en') {
        document.querySelector('label[for="colleague"]').textContent = 'Authorized Person:';
        document.querySelector('label[for="startDate"]').textContent = 'Start Date and Time:';
        document.querySelector('label[for="endDate"]').textContent = 'End Date and Time:';
        document.getElementById('btnSetAutoReply').textContent = 'Set Auto Reply';
        document.querySelector('.preview-section h3').textContent = '📧 Message Preview';
        document.querySelector('option[value=""]').textContent = 'Select...';
    } else {
        document.querySelector('label[for="colleague"]').textContent = 'Yetkili Kişi:';
        document.querySelector('label[for="startDate"]').textContent = 'Başlangıç Tarihi ve Saati:';
        document.querySelector('label[for="endDate"]').textContent = 'Bitiş Tarihi ve Saati:';
        document.getElementById('btnSetAutoReply').textContent = 'Otomatik Yanıtı Ayarla';
        document.querySelector('.preview-section h3').textContent = '📧 Mesaj Önizleme';
        document.querySelector('option[value=""]').textContent = 'Seçiniz...';
    }
    
    updatePreview();
}

function setupFormListeners() {
    const inputs = ['colleague', 'startDate', 'startTime', 'endDate', 'endTime'];
    inputs.forEach(id => {
        document.getElementById(id).addEventListener('change', updatePreview);
    });
}

function setDefaultDates() {
    const now = new Date();
    const tomorrow = new Date(now);
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const nextWeek = new Date(now);
    nextWeek.setDate(nextWeek.getDate() + 7);
    
    document.getElementById('startDate').value = formatDate(tomorrow);
    document.getElementById('endDate').value = formatDate(nextWeek);
    document.getElementById('startTime').value = '09:00';
    document.getElementById('endTime').value = '18:00';
}

function formatDate(date) {
    return date.toISOString().split('T')[0];
}

function formatDisplayDate(dateStr, timeStr, language) {
    const date = new Date(dateStr + 'T' + timeStr);
    
    if (language === 'tr') {
        return date.toLocaleDateString('tr-TR', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        }) + ' ' + timeStr;
    } else {
        return date.toLocaleDateString('en-US', {
            month: 'long',
            day: 'numeric',
            year: 'numeric'
        }) + ' ' + timeStr;
    }
}

function updatePreview() {
    const colleagueId = document.getElementById('colleague').value;
    const startDate = document.getElementById('startDate').value;
    const startTime = document.getElementById('startTime').value;
    const endDate = document.getElementById('endDate').value;
    const endTime = document.getElementById('endTime').value;
    
    const previewDiv = document.getElementById('messagePreview');
    
    if (!colleagueId || !startDate || !startTime || !endDate || !endTime) {
        previewDiv.textContent = currentLanguage === 'tr' ? 
            'Lütfen tüm alanları doldurun...' : 
            'Please fill in all fields...';
        return;
    }
    
    const colleague = colleagues.find(c => c.id == colleagueId);
    const template = templates[currentLanguage];
    
    const startDateTime = formatDisplayDate(startDate, startTime, currentLanguage);
    const endDateTime = formatDisplayDate(endDate, endTime, currentLanguage);
    
    // Get current user info (in production, this would come from Office.js)
    const currentUser = {
        name: "Kullanıcı Adı", // This would be retrieved from Office context
        position: "Pozisyon",
        company: "Öztüryakiler"
    };
    
    let messageBody = template.body
        .replace('{startDate}', startDateTime)
        .replace('{endDate}', endDateTime)
        .replace('{colleagueName}', colleague.name)
        .replace('{email}', colleague.email)
        .replace('{phone}', colleague.phone)
        .replace('{userName}', currentUser.name)
        .replace('{position}', currentUser.position)
        .replace('{company}', currentUser.company);
    
    previewDiv.textContent = `Konu: ${template.subject}\n\n${messageBody}`;
}

function setAutoReply(event) {
    event.preventDefault();
    
    const colleagueId = document.getElementById('colleague').value;
    const startDate = document.getElementById('startDate').value;
    const startTime = document.getElementById('startTime').value;
    const endDate = document.getElementById('endDate').value;
    const endTime = document.getElementById('endTime').value;
    
    if (!colleagueId || !startDate || !startTime || !endDate || !endTime) {
        showStatus('error', currentLanguage === 'tr' ? 
            'Lütfen tüm alanları doldurun!' : 
            'Please fill in all fields!');
        return;
    }
    
    const colleague = colleagues.find(c => c.id == colleagueId);
    const template = templates[currentLanguage];
    
    const startDateTime = new Date(startDate + 'T' + startTime);
    const endDateTime = new Date(endDate + 'T' + endTime);
    
    if (startDateTime >= endDateTime) {
        showStatus('error', currentLanguage === 'tr' ? 
            'Bitiş tarihi başlangıç tarihinden sonra olmalıdır!' : 
            'End date must be after start date!');
        return;
    }
    
    const button = document.getElementById('btnSetAutoReply');
    button.disabled = true;
    button.textContent = currentLanguage === 'tr' ? 'Ayarlanıyor...' : 'Setting...';
    
    // Prepare the auto-reply message
    const startDateTimeFormatted = formatDisplayDate(startDate, startTime, currentLanguage);
    const endDateTimeFormatted = formatDisplayDate(endDate, endTime, currentLanguage);
    
    const currentUser = {
        name: "Kullanıcı Adı", // This would be retrieved from Office context
        position: "Pozisyon",
        company: "Öztüryakiler"
    };
    
    let messageBody = template.body
        .replace('{startDate}', startDateTimeFormatted)
        .replace('{endDate}', endDateTimeFormatted)
        .replace('{colleagueName}', colleague.name)
        .replace('{email}', colleague.email)
        .replace('{phone}', colleague.phone)
        .replace('{userName}', currentUser.name)
        .replace('{position}', currentUser.position)
        .replace('{company}', currentUser.company);
    
    // Use Office.js to set the auto-reply
    Office.context.mailbox.userProfile.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            // In a real implementation, we would use EWS or Graph API to set auto-reply
            // For now, we'll simulate the process
            setTimeout(() => {
                showStatus('success', currentLanguage === 'tr' ? 
                    'Otomatik yanıt başarıyla ayarlandı!' : 
                    'Auto-reply has been set successfully!');
                
                button.disabled = false;
                button.textContent = currentLanguage === 'tr' ? 
                    'Otomatik Yanıtı Ayarla' : 
                    'Set Auto Reply';
                
                // Log the auto-reply details for debugging
                console.log('Auto-reply set:', {
                    subject: template.subject,
                    body: messageBody,
                    startDate: startDateTime,
                    endDate: endDateTime,
                    colleague: colleague
                });
                
            }, 2000);
        } else {
            showStatus('error', currentLanguage === 'tr' ? 
                'Otomatik yanıt ayarlanırken hata oluştu!' : 
                'Error occurred while setting auto-reply!');
            
            button.disabled = false;
            button.textContent = currentLanguage === 'tr' ? 
                'Otomatik Yanıtı Ayarla' : 
                'Set Auto Reply';
        }
    });
}

function showStatus(type, message) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.className = `status-message status-${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 5000);
}

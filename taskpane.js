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
let colleagues = [];

// Message template (both Turkish and English)
const messageTemplate = {
    subject: "Otomatik Yanıt: Yıllık İzin / Automatic Reply: Annual Leave",
    body: `Sayın Yetkili,

E-postanız için teşekkür ederim. {startDate} – {endDate} tarihleri arasında yıllık izinde olacağım ve bu süre içinde e-postalarınıza yanıt veremeyeceğim.

Acil konularınız için {colleagueName} ile {email} veya {phone} üzerinden iletişime geçebilirsiniz.

Anlayışınız için teşekkür eder, iyi çalışmalar dilerim.

Saygılarımla,
{userName}
{position}
{company}

---

Dear Sir/Madam,

Thank you for your email. I will be out of the office on annual leave from {startDate} to {endDate}, and will not be able to respond to your message during this period.

For urgent matters, please contact {colleagueName} at {email} or {phone}.

Thank you for your understanding.

Kind regards,
{userName}
{position}
{company}`
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
    },
    {
        id: 6,
        name: "Zeynep Arslan",
        email: "zeynep.arslan@ozturyakiler.com.tr",
        phone: "+90 212 555 0106",
        department: "Hukuk"
    },
    {
        id: 7,
        name: "Murat Çelik",
        email: "murat.celik@ozturyakiler.com.tr",
        phone: "+90 212 555 0107",
        department: "Finans"
    },
    {
        id: 8,
        name: "Elif Koç",
        email: "elif.koc@ozturyakiler.com.tr",
        phone: "+90 212 555 0108",
        department: "Operasyon"
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


function setupFormListeners() {
    const inputs = ['colleague', 'startDate', 'startTime', 'endDate', 'endTime'];
    inputs.forEach(id => {
        document.getElementById(id).addEventListener('change', updatePreview);
    });
}

function setDefaultDates() {
    const today = new Date();
    const nextWeek = new Date(today);
    nextWeek.setDate(nextWeek.getDate() + 7);
    
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    
    // Set minimum dates
    startDateInput.min = formatDate(today);
    endDateInput.min = formatDate(today);
    
    // Set default values
    startDateInput.value = formatDate(today);
    endDateInput.value = formatDate(nextWeek);
    
    // Add event listener to update end date minimum when start date changes
    startDateInput.addEventListener('change', function() {
        const startDate = new Date(this.value);
        endDateInput.min = formatDate(startDate);
        
        // If end date is before start date, update it
        if (endDateInput.value && new Date(endDateInput.value) < startDate) {
            endDateInput.value = formatDate(startDate);
        }
    });
}

function formatDate(date) {
    return date.toISOString().split('T')[0];
}

function formatDisplayDate(dateStr, timeStr) {
    const date = new Date(dateStr + 'T' + timeStr);
    
    return date.toLocaleDateString('tr-TR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    }) + ' ' + timeStr;
}

function updatePreview() {
    const colleagueId = document.getElementById('colleague').value;
    const startDate = document.getElementById('startDate').value;
    const startTime = document.getElementById('startTime').value;
    const endDate = document.getElementById('endDate').value;
    const endTime = document.getElementById('endTime').value;
    
    const previewDiv = document.getElementById('messagePreview');
    
    if (!colleagueId || !startDate || !startTime || !endDate || !endTime) {
        previewDiv.textContent = 'Lütfen tüm alanları doldurun...';
        return;
    }
    
    const colleague = colleagues.find(c => c.id == colleagueId);
    
    const startDateTime = formatDisplayDate(startDate, startTime);
    const endDateTime = formatDisplayDate(endDate, endTime);
    
    // Get current user info (in production, this would come from Office.js)
    const currentUser = {
        name: "Kullanıcı Adı", // This would be retrieved from Office context
        position: "Pozisyon",
        company: "Öztüryakiler"
    };
    
    let messageBody = messageTemplate.body
        .replace('{startDate}', startDateTime)
        .replace('{endDate}', endDateTime)
        .replace('{colleagueName}', colleague.name)
        .replace('{email}', colleague.email)
        .replace('{phone}', colleague.phone)
        .replace('{userName}', currentUser.name)
        .replace('{position}', currentUser.position)
        .replace('{company}', currentUser.company);
    
    previewDiv.textContent = `Konu: ${messageTemplate.subject}\n\n${messageBody}`;
}

function setAutoReply(event) {
    event.preventDefault();
    
    const colleagueId = document.getElementById('colleague').value;
    const startDate = document.getElementById('startDate').value;
    const startTime = document.getElementById('startTime').value;
    const endDate = document.getElementById('endDate').value;
    const endTime = document.getElementById('endTime').value;
    
    if (!colleagueId || !startDate || !startTime || !endDate || !endTime) {
        showStatus('error', 'Lütfen tüm alanları doldurun!');
        return;
    }
    
    const colleague = colleagues.find(c => c.id == colleagueId);
    
    const startDateTime = new Date(startDate + 'T' + startTime);
    const endDateTime = new Date(endDate + 'T' + endTime);
    
    if (startDateTime >= endDateTime) {
        showStatus('error', 'Bitiş tarihi başlangıç tarihinden sonra olmalıdır!');
        return;
    }
    
    const button = document.getElementById('btnSetAutoReply');
    button.disabled = true;
    button.textContent = 'Ayarlanıyor...';
    
    // Prepare the auto-reply message
    const startDateTimeFormatted = formatDisplayDate(startDate, startTime);
    const endDateTimeFormatted = formatDisplayDate(endDate, endTime);
    
    const currentUser = {
        name: "Kullanıcı Adı", // This would be retrieved from Office context
        position: "Pozisyon",
        company: "Öztüryakiler"
    };
    
    let messageBody = messageTemplate.body
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
                showStatus('success', 'Otomatik yanıt başarıyla ayarlandı!');
                
                button.disabled = false;
                button.textContent = 'Otomatik Yanıtı Ayarla';
                
                // Log the auto-reply details for debugging
                console.log('Auto-reply set:', {
                    subject: messageTemplate.subject,
                    body: messageBody,
                    startDate: startDateTime,
                    endDate: endDateTime,
                    colleague: colleague
                });
                
            }, 2000);
        } else {
            showStatus('error', 'Otomatik yanıt ayarlanırken hata oluştu!');
            
            button.disabled = false;
            button.textContent = 'Otomatik Yanıtı Ayarla';
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

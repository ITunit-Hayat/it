(function () {
'use strict';

/* ── URL الـ Google Apps Script ── */
var SURL = window.IT_CONFIG ? window.IT_CONFIG.SURL : '';

/* ══════════════════════════════════════════════
   شريط التنقل — شفافية عند الأعلى
══════════════════════════════════════════════ */
var navEl = document.getElementById('nav');
navEl.classList.add('transparent');
window.addEventListener('scroll', function () {
  var hero  = document.getElementById('hero');
  var limit = hero ? hero.offsetHeight - 80 : 200;
  if (window.scrollY < limit) navEl.classList.add('transparent');
  else                         navEl.classList.remove('transparent');

  /* زر الأعلى */
  var topBtn = document.getElementById('back-top');
  if (window.scrollY > 400) topBtn.classList.add('show');
  else                       topBtn.classList.remove('show');
});

/* ══════════════════════════════════════════════
   همبرغر موبايل
══════════════════════════════════════════════ */
var hbg      = document.getElementById('hamburger');
var mobMenu  = document.getElementById('mobile-menu');
hbg.addEventListener('click', function () {
  hbg.classList.toggle('open');
  mobMenu.classList.toggle('open');
});
document.querySelectorAll('.mobile-link').forEach(function(a) {
  a.addEventListener('click', function () {
    hbg.classList.remove('open');
    mobMenu.classList.remove('open');
  });
});

/* ══════════════════════════════════════════════
   Scroll-Reveal
══════════════════════════════════════════════ */
if ('IntersectionObserver' in window) {
  var observer = new IntersectionObserver(function (entries) {
    entries.forEach(function (entry) {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.1 });
  document.querySelectorAll('.reveal').forEach(function (el) { observer.observe(el); });
}

/* ══════════════════════════════════════════════
   انيميشن الأرقام
══════════════════════════════════════════════ */
function animNum(id, val, suffix) {
  var el   = document.getElementById(id);
  if (!el) return;
  el.innerHTML = '';
  var c    = 0;
  var s    = suffix || '';
  var step = Math.max(val / 60, 0.5);
  var tm   = setInterval(function () {
    c = Math.min(c + step, val);
    el.textContent = Math.floor(c) + s;
    if (c >= val) clearInterval(tm);
  }, 25);
}

/* ══════════════════════════════════════════════
   تحميل الإحصائيات
══════════════════════════════════════════════ */
function loadStats() {
  fetch(SURL + '?action=stats')
    .then(function (r) { return r.json(); })
    .then(function (d) {
      var tot  = d.total      || 0;
      var fix  = d.fixed      || 0;
      var opn  = d.open       || 0;
      var proc = d.processing || 0;
      var pct  = tot > 0 ? Math.round(fix / tot * 100) : 0;
      animNum('st-total',      tot,  '');
      animNum('st-fixed',      pct,  '%');
      animNum('st-open',       opn,  '');
      animNum('st-processing', proc, '');
    })
    .catch(function () {
      ['st-total','st-fixed','st-open','st-processing'].forEach(function (id) {
        var el = document.getElementById(id);
        if (el) el.textContent = '—';
      });
    });
}

/* ══════════════════════════════════════════════
   Toast
══════════════════════════════════════════════ */
var toastTimer = null;
function showToast(type, title, sub, tags) {
  var el = document.getElementById('toast');
  document.getElementById('t-icon').textContent  = type === 'ok' ? '✅' : '⚠️';
  document.getElementById('t-title').textContent = title;
  document.getElementById('t-sub').textContent   = sub;
  var tagsEl = document.getElementById('t-tags');
  tagsEl.innerHTML = '';
  (tags || []).forEach(function (t) {
    var s = document.createElement('span');
    s.className = 'toast-tag'; s.textContent = t;
    tagsEl.appendChild(s);
  });
  el.className = 'show ' + type;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.className = type; }, 7000);
}

window.closeToast = function () {
  var el = document.getElementById('toast');
  el.className = el.className.replace('show', '').trim();
};

/* ══════════════════════════════════════════════
   عداد الأحرف
══════════════════════════════════════════════ */
document.getElementById('f-details').addEventListener('input', function () {
  document.getElementById('char-count').textContent = this.value.length;
});

/* ══════════════════════════════════════════════
   التحقق من الحقول
══════════════════════════════════════════════ */
function showError(inputId, errId, show) {
  var inp = document.getElementById(inputId);
  var err = document.getElementById(errId);
  if (!inp || !err) return;
  if (show) { inp.classList.add('error'); err.style.display = 'block'; }
  else       { inp.classList.remove('error'); err.style.display = 'none'; }
}

function validate() {
  var name    = trim('f-name');
  var phone   = trim('f-phone');
  var email   = trim('f-email');
  var branch  = trim('f-branch');
  var admin   = trim('f-admin');
  var details = trim('f-details');
  var ok = true;

  showError('f-name',    'err-name',    name.length < 2);
  showError('f-phone',   'err-phone',   phone.replace(/\D/g,'').length < 8);
  showError('f-email',   'err-email',   email.indexOf('@') < 0);
  showError('f-branch',  'err-branch',  !branch && admin !== 'الإدارة العامة' && admin !== 'معهد الحياة القديم');
  showError('f-admin',   'err-admin',   !admin);
  showError('f-details', 'err-details', details.length < 5);

  if (name.length < 2  || phone.replace(/\D/g,'').length < 8 ||
      email.indexOf('@') < 0 || (!branch && admin !== 'الإدارة العامة' && admin !== 'معهد الحياة القديم') || !admin || details.length < 5)
    ok = false;

  return ok;
}

function trim(id) {
  var el = document.getElementById(id);
  return el ? (el.value || '').trim() : '';
}

/* ── إزالة حالة الخطأ عند الكتابة ── */
['f-name','f-phone','f-email','f-branch','f-admin','f-details'].forEach(function (id) {
  var el = document.getElementById(id);
  if (!el) return;
  el.addEventListener('input', function () {
    el.classList.remove('error');
    var errId = id.replace('f-', 'err-');
    var err = document.getElementById(errId);
    if (err) err.style.display = 'none';
  });
});

/* ══════════════════════════════════════════════
   إرسال النموذج
══════════════════════════════════════════════ */
document.getElementById('submit-btn').addEventListener('click', sendForm);

function sendForm() {
  if (!validate()) {
    showToast('er', 'يوجد خطأ في البيانات', 'يرجى مراجعة الحقول المُعلَّمة باللون الأحمر', []);
    return;
  }

  var branch   = trim('f-branch');
  var adminVal = trim('f-admin');
  var subAdmin = document.getElementById('f-admin-sub') ? (document.getElementById('f-admin-sub').value || '') : '';
  var dept     = (branch && adminVal) ? branch + ' - ' + adminVal + (subAdmin ? ' / ' + subAdmin : '') : (branch || adminVal);
  var phone    = '+213' + trim('f-phone').replace(/^0+/, '');
  var tkt      = 'TKT-' + Date.now().toString().slice(-6);

  var payload = JSON.stringify({
    name:     trim('f-name'),
    empId:    trim('f-empid'),
    phone:    phone,
    email:    trim('f-email'),
    dept:     dept,
    type:     document.getElementById('f-type').value,
    priority: document.getElementById('f-priority').value,
    details:  trim('f-details'),
    ticket:   tkt
  });

  var btn = document.getElementById('submit-btn');
  btn.disabled = true;
  btn.classList.add('btn-loading');

  fetch(SURL, {
    method: 'POST',
    mode:   'no-cors',
    headers:{ 'Content-Type': 'text/plain;charset=UTF-8' },
    body:   payload
  })
  .then(function () {
    showToast('ok',
      '✅ تم إرسال طلب التدخل!',
      'رقم التتبع: ' + tkt + ' · ستصلك رسالة تأكيد على بريدك الإلكتروني',
      ['💬 تيليغرام', '📊 Sheets', '📧 إيميل', '💬 واتساب']
    );
    resetForm();
  })
  .catch(function () {
    showToast('er', 'تعذّر الإرسال', 'تحقق من اتصالك أو تواصل معنا: 0779149963', []);
  })
  .finally(function () {
    btn.disabled = false;
    btn.classList.remove('btn-loading');
  });
}

function resetForm() {
  ['f-name','f-empid','f-phone','f-email','f-details'].forEach(function (id) {
    var el = document.getElementById(id);
    if (el) el.value = '';
  });
  document.getElementById('f-branch').value   = '';
  document.getElementById('f-admin').value    = '';
  document.getElementById('f-type').value     = 'دعم تقني عاجل';
  document.getElementById('f-priority').value = 'عادي';
  document.getElementById('char-count').textContent = '0';
}

/* ══════════════════════════════════════════════
   تتبع التذكرة
══════════════════════════════════════════════ */
window.lookupTicket = function () {
  var tkt = document.getElementById('tracker-input').value.trim().toUpperCase();
  if (!tkt) {
    showToast('er', 'أدخل رقم التذكرة', 'مثال: TKT-0001', []);
    return;
  }

  var resultEl = document.getElementById('tracker-result');
  var btn      = document.getElementById('tracker-btn');
  resultEl.style.display = 'block';
  resultEl.innerHTML = '<div style="text-align:center;padding:30px;color:var(--mt)">🔍 جارٍ البحث...</div>';
  btn.disabled = true;

  fetch(SURL + '?action=lookup&ticket=' + encodeURIComponent(tkt))
    .then(function (r) { return r.json(); })
    .then(function (d) {
      if (d.found) {
        var sClass = getStatusClass(d.status);
        resultEl.innerHTML = [
          '<div class="tracker-card">',
          '<div class="ticket-id">🎫 ' + d.ticket + '</div>',
          buildTrackerRow('الحالة',         '<span class="status-badge ' + sClass + '">' + d.status + '</span>'),
          buildTrackerRow('الاسم',           escHtml(d.name    || '—')),
          buildTrackerRow('نوع الطلب',       escHtml(d.type    || '—')),
          buildTrackerRow('الفرع',           escHtml(d.dept    || '—')),
          buildTrackerRow('الأولوية',        escHtml(d.priority|| '—')),
          buildTrackerRow('تاريخ الإرسال',   escHtml(d.date    || '—')),
          d.closed ? buildTrackerRow('تاريخ الإغلاق', escHtml(d.closed)) : '',
          d.notes  ? buildTrackerRow('ملاحظة الفريق', escHtml(d.notes))  : '',
          '</div>'
        ].join('');
      } else {
        resultEl.innerHTML = '<div class="tracker-not-found"><div style="font-size:48px;margin-bottom:12px">🔍</div><p>لم يُوجد طلب بالرقم <strong style="color:var(--gl)">' + escHtml(tkt) + '</strong></p><p style="font-size:13px;margin-top:8px">تأكد من الرقم أو تواصل مع الفريق على 0779149963</p></div>';
      }
    })
    .catch(function () {
      resultEl.innerHTML = '<div class="tracker-not-found">⚠️ تعذّر الاتصال بالخادم. تحقق من اتصالك.</div>';
    })
    .finally(function () { btn.disabled = false; });
};

function buildTrackerRow(label, value) {
  return '<div class="tracker-row"><div class="tracker-row-label">' + label + '</div><div class="tracker-row-value">' + value + '</div></div>';
}
function getStatusClass(s) {
  s = s || '';
  if (s.indexOf('تم الإصلاح')   >= 0) return 'status-fixed';
  if (s.indexOf('مرفوض')        >= 0) return 'status-rejected';
  if (s.indexOf('قيد المعالجة') >= 0) return 'status-processing';
  if (s.indexOf('عاجل')         >= 0) return 'status-urgent';
  return 'status-open';
}
function escHtml(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}

/* ══════════════════════════════════════════════
   FAQ أكورديون
══════════════════════════════════════════════ */
window.toggleFaq = function (btn) {
  var ans     = btn.nextElementSibling;
  var isOpen  = btn.classList.contains('open');
  /* إغلاق الكل */
  document.querySelectorAll('.faq-q.open').forEach(function (b) {
    b.classList.remove('open');
    b.nextElementSibling.classList.remove('open');
  });
  if (!isOpen) {
    btn.classList.add('open');
    ans.classList.add('open');
  }
};

/* ══════════════════════════════════════════════
   زر الرجوع للأعلى
══════════════════════════════════════════════ */
window.scrollToTop = function () {
  window.scrollTo({ top: 0, behavior: 'smooth' });
};

/* ══════════════════════════════════════════════
   تهيئة عند التحميل
══════════════════════════════════════════════ */
window.addEventListener('load', function () {
    var adminSel = document.getElementById('f-admin');
    if (adminSel) {
      adminSel.addEventListener('change', function () {
        var wrap   = document.getElementById('sub-admin-wrap');
        var branch = document.getElementById('f-branch');
        if (this.value === 'الإدارة العامة' || this.value === 'معهد الحياة القديم') {
          wrap.style.display = 'block';
          branch.value = ''; branch.disabled = true;
          branch.style.opacity = '0.4'; branch.style.cursor = 'not-allowed';
        } else {
          wrap.style.display = 'none';
          document.getElementById('f-admin-sub').value = '';
          branch.disabled = false; branch.style.opacity = '1'; branch.style.cursor = 'pointer';
        }
      });
    }
  loadStats();
  /* Enter في حقل تتبع التذكرة */
  document.getElementById('tracker-input').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') window.lookupTicket();
  });
  /* تنسيق رقم التذكرة تلقائياً */
  document.getElementById('tracker-input').addEventListener('input', function () {
    this.value = this.value.toUpperCase().replace(/[^A-Z0-9\-]/g, '');
  });
});

})();
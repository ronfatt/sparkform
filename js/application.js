/**
 * Spark One 申请中心控制器 - 4大官方 A4 表单
 */

(function () {
  'use strict';

  let currentLang = 'zh-CN';
  let activeFormId = 'office-subsidy';
  let isCanvasDrawn = false;
  let canvas, ctx;

  const formMeta = {
    'office-subsidy': '办公室补贴申请表',
    'lecturer-subsidy': '讲师补贴申请表',
    'event-apply': 'SPARK UNION CAPITAL 活动申请表',
    'pilot-grant': '市场领航扶持账号额度申请表'
  };

  // 1. 切换表单
  function switchFormView(formId) {
    if (!formId) return;
    if (formId === 'hub') {
      showDirectoryView();
      return;
    }
    activeFormId = formId;

    const viewDirectory = document.getElementById('viewDirectory');
    const viewFormWorkspace = document.getElementById('viewFormWorkspace');

    if (viewDirectory) viewDirectory.style.display = 'none';
    if (viewFormWorkspace) viewFormWorkspace.style.display = 'block';

    document.querySelectorAll('.form-wrapper').forEach(el => el.style.display = 'none');
    const target = document.getElementById(`formSection-${formId}`);
    if (target) target.style.display = 'flex';

    const formQuickSelect = document.getElementById('formQuickSelect');
    const formInnerSelect = document.getElementById('formInnerSelect');
    if (formQuickSelect) formQuickSelect.value = formId;
    if (formInnerSelect) formInnerSelect.value = formId;

    const activeFormTitle = document.getElementById('activeFormTitle');
    if (activeFormTitle) {
      const dict = (window.APPLICATION_I18N && window.APPLICATION_I18N[currentLang]) 
                || (window.APPLICATION_I18N && window.APPLICATION_I18N['zh-CN']) 
                || {};
      const formKeyMap = {
        'office-subsidy': 'off_title',
        'lecturer-subsidy': 'lec_title',
        'event-apply': 'act_main_title',
        'pilot-grant': 'pilot_title'
      };
      const key = formKeyMap[formId];
      activeFormTitle.textContent = (key && dict[key]) ? dict[key] : (formMeta[formId] || formId);
    }

    // 默认给各表单日期赋今天的值
    const today = new Date().toISOString().split('T')[0];
    const isZh = (currentLang === 'zh-CN' || currentLang === 'zh-TW');
    const todayStr = isZh ? `${new Date().getFullYear()}年${new Date().getMonth()+1}月${new Date().getDate()}日` : today;
    ['off_date', 'lec_date', 'act_date', 'act_sign_date', 'sec7_date'].forEach(id => {
      const el = document.getElementById(id);
      if (el && !el.value) el.value = el.type === 'date' ? today : todayStr;
    });

    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function showDirectoryView() {
    const viewDirectory = document.getElementById('viewDirectory');
    const viewFormWorkspace = document.getElementById('viewFormWorkspace');
    if (viewDirectory) viewDirectory.style.display = 'block';
    if (viewFormWorkspace) viewFormWorkspace.style.display = 'none';
    const formQuickSelect = document.getElementById('formQuickSelect');
    if (formQuickSelect) formQuickSelect.value = 'hub';
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function showToast(msg) {
    const toastMsg = document.getElementById('toastMsg');
    if (!toastMsg) return;
    toastMsg.textContent = msg;
    toastMsg.classList.add('show');
    setTimeout(() => toastMsg.classList.remove('show'), 2600);
  }

  // 2. 签名板初始化
  function initSignatureCanvas() {
    try {
      canvas = document.getElementById('signatureCanvas');
      if (!canvas) return;
      ctx = canvas.getContext('2d');
      if (!ctx) return;

      const rect = canvas.getBoundingClientRect();
      const dpr = window.devicePixelRatio || 1;
      const w = rect.width > 0 ? rect.width : 280;
      const h = rect.height > 0 ? rect.height : 80;
      canvas.width = w * dpr;
      canvas.height = h * dpr;
      ctx.scale(dpr, dpr);

      ctx.lineWidth = 2.2;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = '#0f172a';

      let drawing = false;
      function getPos(e) {
        const b = canvas.getBoundingClientRect();
        const cx = e.touches ? e.touches[0].clientX : e.clientX;
        const cy = e.touches ? e.touches[0].clientY : e.clientY;
        return { x: cx - b.left, y: cy - b.top };
      }

      canvas.addEventListener('mousedown', e => {
        drawing = true;
        isCanvasDrawn = true;
        const p = getPos(e);
        ctx.beginPath();
        ctx.moveTo(p.x, p.y);
      });
      canvas.addEventListener('mousemove', e => {
        if (!drawing) return;
        const p = getPos(e);
        ctx.lineTo(p.x, p.y);
        ctx.stroke();
      });
      window.addEventListener('mouseup', () => { drawing = false; if (ctx) ctx.closePath(); });

      canvas.addEventListener('touchstart', e => {
        drawing = true;
        isCanvasDrawn = true;
        const p = getPos(e);
        ctx.beginPath();
        ctx.moveTo(p.x, p.y);
        e.preventDefault();
      }, { passive: false });
      canvas.addEventListener('touchmove', e => {
        if (!drawing) return;
        const p = getPos(e);
        ctx.lineTo(p.x, p.y);
        ctx.stroke();
        e.preventDefault();
      }, { passive: false });
      window.addEventListener('touchend', () => { drawing = false; });
    } catch (e) {
      console.warn('initSignatureCanvas error:', e);
    }
  }

  function clearSignature() {
    if (!ctx || !canvas) return;
    try {
      const rect = canvas.getBoundingClientRect();
      const dpr = window.devicePixelRatio || 1;
      ctx.clearRect(0, 0, rect.width * dpr, rect.height * dpr);
      isCanvasDrawn = false;
    } catch (e) {}
  }

  // 3. 语言切换 (全局暴露)
  function updateLanguage(lang) {
    if (!lang) return;
    currentLang = lang;
    window.currentAppLang = lang;

    try {
      localStorage.setItem('spark_form_lang', lang);
    } catch (e) {}

    document.documentElement.setAttribute('dir', lang === 'ar' ? 'rtl' : 'ltr');
    document.documentElement.setAttribute('lang', lang);

    // 同步语言下拉选择框
    const ls = document.getElementById('langSelect');
    if (ls && ls.value !== lang) {
      ls.value = lang;
    }

    const dict = (window.APPLICATION_I18N && window.APPLICATION_I18N[lang]) 
              || (window.APPLICATION_I18N && window.APPLICATION_I18N['zh-CN']) 
              || {};

    // 翻译页面所有标记了 data-i18n 的元素
    document.querySelectorAll('[data-i18n]').forEach(el => {
      const key = el.getAttribute('data-i18n');
      if (dict && dict[key] !== undefined) {
        el.textContent = dict[key];
      }
    });

    // 翻译页面所有标记了 data-i18n-ph 的输入框 placeholder
    document.querySelectorAll('[data-i18n-ph]').forEach(el => {
      const key = el.getAttribute('data-i18n-ph');
      if (dict && dict[key] !== undefined) {
        el.setAttribute('placeholder', dict[key]);
      }
    });

    // 动态同步更新各表单日期输入框格式（防止非中文显示中文年月日）
    const isZhLang = (lang === 'zh-CN' || lang === 'zh-TW');
    const d = new Date();
    const curTodayStr = isZhLang 
      ? `${d.getFullYear()}年${String(d.getMonth()+1).padStart(2,'0')}月${String(d.getDate()).padStart(2,'0')}日` 
      : d.toISOString().split('T')[0];
    ['off_date', 'lec_date', 'act_date'].forEach(id => {
      const el = document.getElementById(id);
      if (el) {
        if (!el.value || /^\d{4}[年\-\/]\d{1,2}[月\-\/]\d{1,2}[日]?$/.test(el.value.trim())) {
          el.value = curTodayStr;
        }
      }
    });

    // 翻译面包屑当前表单标题
    const activeFormTitle = document.getElementById('activeFormTitle');
    if (activeFormTitle) {
      const formKeyMap = {
        'office-subsidy': 'off_title',
        'lecturer-subsidy': 'lec_title',
        'event-apply': 'act_main_title',
        'pilot-grant': 'pilot_title'
      };
      const key = formKeyMap[activeFormId];
      if (key && dict[key]) {
        activeFormTitle.textContent = dict[key];
      }
    }

    // 动态同步更新网页 Title 和 Meta / OG / Twitter 描述
    if (dict && dict.meta_title) {
      document.title = dict.meta_title;
      const ogTitle = document.querySelector('meta[property="og:title"]');
      if (ogTitle) ogTitle.setAttribute('content', dict.meta_title);
      const twTitle = document.querySelector('meta[name="twitter:title"]');
      if (twTitle) twTitle.setAttribute('content', dict.meta_title);
    }
    if (dict && dict.meta_desc) {
      const metaDesc = document.querySelector('meta[name="description"]');
      if (metaDesc) metaDesc.setAttribute('content', dict.meta_desc);
      const ogDesc = document.querySelector('meta[property="og:description"]');
      if (ogDesc) ogDesc.setAttribute('content', dict.meta_desc);
      const twDesc = document.querySelector('meta[name="twitter:description"]');
      if (twDesc) twDesc.setAttribute('content', dict.meta_desc);
    }

    const langToastMap = {
      'zh-CN': '已切换至 简体中文',
      'zh-TW': '已切換至 繁體中文',
      'en': 'Switched to English',
      'ko': '한국어로 전환되었습니다',
      'ja': '日本語に切り替えました',
      'th': 'เปลี่ยนเป็นภาษาไทยแล้ว',
      'vi': 'Đã chuyển sang Tiếng Việt',
      'id': 'Beralih ke Bahasa Indonesia',
      'ms': 'Ditukar ke Bahasa Melayu',
      'es': 'Cambiado a Español',
      'pt': 'Alterado para Português',
      'ar': 'تم التحويل إلى اللغة العربية'
    };
    showToast(langToastMap[lang] || `Language: ${lang}`);
  }

  // 4. 收集当前表单数据
  function collectFormData() {
    const data = {};
    const wrap = document.getElementById(`formSection-${activeFormId}`);
    if (!wrap) return data;

    wrap.querySelectorAll('input, textarea, select').forEach(el => {
      if (el.id) {
        data[el.id] = el.value;
      }
    });
    return data;
  }

  // 5. 填入示范数据
  // 5. 填入示范数据
  function loadDemoData() {
    const isZh = currentLang === 'zh-CN' || currentLang === 'zh-TW';
    if (activeFormId === 'office-subsidy') {
      const d = {
        off_date: isZh ? '2026年10月15日' : '2026-10-15',
        off_lead_name: 'David Tan (陈志远)',
        off_lead_contact: '+65 9182 3746 / @david_spark',
        off_lead_city: 'Singapore (新加坡)',
        off_lead_team: 'Apex APAC Team',
        off_address: 'Marina Bay Financial Centre (MBFC Tower 2) #18-06, Singapore',
        off_area: '150',
        off_capacity: '35',
        off_rent: '4800',
        off_rent_curr: 'USD',
        off_lease_period: isZh ? '2026年11月01日 至 2027年10月31日' : '2026-11-01 - 2027-10-31',
        off_start_date: '2026-11-01',
        off_sub_amt: '2000',
        off_sub_curr: 'USD',
        off_sub_period: isZh ? '2026年11月01日 至 2027年04月30日' : '2026-11-01 - 2027-04-30',
        off_comm_name: 'Spark One APAC Core Wealth Club (Telegram)',
        off_comm_link: 'https://t.me/Spark_Apex_APAC',
        off_comm_admins: 'David Tan, Evelyn Lim',
        off_comm_members: '520',
        off_comm_active: '120+ active daily, 350+ weekly webinar participants',
        off_comm_lang: 'English / Chinese / Malay',
        off_comm_daily: '09:00 Market Morning Call; 14:00 Arbitrage Analysis; 21:00 Daily Profit Report',
        off_comm_weekly: 'Tuesday online Zoom workshop; Saturday offline VIP office salon',
        off_comm_feedback: '1-on-1 advisor channel, 30-min response SLA',
        off_kpi_period: isZh ? '2026年11月01日 至 2026年11月30日' : '2026-11-01 - 2026-11-30',
        off_kpi_goal: 'Direct and team quantitative custody exceeding 250,000 USD'
      };
      Object.keys(d).forEach(k => {
        const el = document.getElementById(k);
        if (el) el.value = d[k];
      });
    } else if (activeFormId === 'lecturer-subsidy') {
      const d = {
        lec_date: isZh ? '2026年10月15日' : '2026-10-15',
        lec_name: 'Raymond Zhang (张峰)',
        lec_account: 'SPK-882910 / Senior Quantitative Lecturer',
        lec_city: isZh ? '中国深圳 / 香港' : 'Shenzhen / Hong Kong',
        lec_contact: '+86 138 2888 9999 / raymond@spark1.cc',
        lec_lang: isZh ? '中文普通话、粤语、商务英语' : 'English / Chinese Mandarin',
        lec_team: isZh ? '大湾区国际服务中心' : 'GBA International Center',
        lec_ref: 'David Tan (+65 91823746)',
        lec_exp: '8 years hedge fund & crypto quant training experience, 50+ HNWI keynotes',
        lec_amt: '2,500 USD / Month',
        lec_audience: 'HNWIs and team leaders, focusing on MSB compliance & steady arbitrage yield logic',
        lec_kpi_period: isZh ? '2026年11月01日 至 2026年11月30日' : '2026-11-01 - 2026-11-30',
        lec_goal: 'Minimum 18 lecture hours/month, assist team in achieving $300k quant expansion',
        lec_cond: 'Student satisfaction ≥ 95%, submit recordings daily',
        lec_src: 'Live trading dashboard and independent settlement logs',
        lec_sign: 'Raymond Zhang',
        lec_ref_sign: 'David Tan'
      };
      Object.keys(d).forEach(k => {
        const el = document.getElementById(k);
        if (el) el.value = d[k];
      });
    } else if (activeFormId === 'event-apply') {
      const d = {
        act_name: isZh ? '2026 Spark One 全球多资产配置与AI量化私享会（新加坡站）' : '2026 Spark One Global Asset Allocation & AI Quant Summit (Singapore)',
        act_date: '2026-11-20',
        act_venue: 'Marina Bay Sands Expo & Convention Centre, Level 4 Grand Ballroom',
        act_app_name: 'Marcus Li (李立明)',
        act_app_phone: '+65 8299 1188',
        act_app_email: 'marcus.li@sparkunioncapital.com',
        act_app_team: 'Singapore Lion City Partner Team',
        act_theme: 'Empowering Wealth: AI Gold Quant & Multi-Asset Steady Hedging Strategies',
        act_b1: '1200',
        act_b2: '600',
        act_b3: '400',
        act_btotal: '2200',
        act_resp1: 'Marcus Li (+65 82991188) Marketing Director',
        act_resp2: 'Raymond Zhang (+86 13828889999) Master Lecturer',
        act_sign: 'Marcus Li',
        act_sign_date: '2026-10-15',
        act_p_num: '48',
        act_reg_num: '22',
        act_inv_num: '15',
        act_pic_num: '85',
        act_summary: isZh ? '全场到会48位高净值客户，现场达成意向资金超38万美元，圆满达成招募考核指标。' : '48 VIP attendees, achieved over $380,000 in custody intent, exceeding recruitment KPI target.'
      };
      Object.keys(d).forEach(k => {
        const el = document.getElementById(k);
        if (el) el.value = d[k];
      });
    } else if (activeFormId === 'pilot-grant') {
      const d = {
        sec1_name: 'David Tan (陈志远)',
        sec1_account: 'SPK-883921',
        sec1_contact: '+65 9182 3746',
        sec1_residence: 'Singapore (Marina Bay)',
        sec1_lang: 'English / Chinese',
        sec1_team: 'Apex APAC Team',
        sec1_job: '10 years quant hedge asset allocation experience',
        sec1_hours: '35',
        sec1_perf_detail: '$280,000 direct quant custody',
        sec2_regions: 'Singapore, Kuala Lumpur, Shenzhen',
        sec2_reason: 'Expand family offices and secure 10x quota conversion within 30 days',
        sec4_target_clients: isZh ? '200人' : '200',
        sec4_target_events: isZh ? '4场' : '4',
        sec7_date: '2026-10-15'
      };
      Object.keys(d).forEach(k => {
        const el = document.getElementById(k);
        if (el) el.value = d[k];
      });
    }

    showToast('已填入专属示范数据！');
  }

  function clearForm() {
    if (!confirm('确定清空当前表格吗？')) return;
    const wrap = document.getElementById(`formSection-${activeFormId}`);
    if (wrap) {
      wrap.querySelectorAll('input, textarea').forEach(el => el.value = '');
    }
    clearSignature();
    showToast('表单已清空');
  }

  // 6. 绑定事件
  function bindEvents() {
    const btnToggleDirectory = document.getElementById('btnToggleDirectory');
    if (btnToggleDirectory) btnToggleDirectory.onclick = showDirectoryView;

    const crumbBackLink = document.getElementById('crumbBackLink');
    if (crumbBackLink) crumbBackLink.onclick = showDirectoryView;

    const formQuickSelect = document.getElementById('formQuickSelect');
    if (formQuickSelect) {
      formQuickSelect.onchange = function () {
        switchFormView(this.value);
      };
    }

    const formInnerSelect = document.getElementById('formInnerSelect');
    if (formInnerSelect) {
      formInnerSelect.onchange = function () {
        switchFormView(this.value);
      };
    }

    const langSelect = document.getElementById('langSelect');
    if (langSelect) {
      langSelect.onchange = function () {
        updateLanguage(this.value);
      };
    }

    // 分类筛选
    document.querySelectorAll('.filter-chip').forEach(chip => {
      chip.onclick = function () {
        document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
        this.classList.add('active');
        const cat = this.getAttribute('data-cat');
        document.querySelectorAll('.form-hub-card').forEach(card => {
          if (cat === 'all' || card.getAttribute('data-category') === cat) {
            card.style.display = 'flex';
          } else {
            card.style.display = 'none';
          }
        });
      };
    });

    const btnFillDemo = document.getElementById('btnFillDemo');
    if (btnFillDemo) btnFillDemo.onclick = loadDemoData;

    const btnClear = document.getElementById('btnClear');
    if (btnClear) btnClear.onclick = clearForm;

    const btnClearSign = document.getElementById('btnClearSign');
    if (btnClearSign) btnClearSign.onclick = clearSignature;

    // 导出 Word (.docx)
    const btnExportWord = document.getElementById('btnExportWord');
    if (btnExportWord) {
      btnExportWord.onclick = async function () {
        const text = this.innerHTML;
        this.innerHTML = '⏳ 生成 Word...';
        this.disabled = true;
        try {
          const data = collectFormData();
          const i18n = (window.APPLICATION_I18N && window.APPLICATION_I18N[currentLang]) 
                    || (window.APPLICATION_I18N && window.APPLICATION_I18N['zh-CN']) 
                    || {};
          const file = await window.ApplicationExporter.exportToWord(activeFormId, data, i18n);
          showToast(`已导出: ${file}`);
        } catch (e) {
          console.error(e);
          alert('导出 Word 失败: ' + e.message);
        } finally {
          this.innerHTML = text;
          this.disabled = false;
        }
      };
    }

    // 导出 PDF (.pdf)
    const btnExportPdf = document.getElementById('btnExportPdf');
    if (btnExportPdf) {
      btnExportPdf.onclick = async function () {
        const text = this.innerHTML;
        this.innerHTML = '⏳ 生成 A4 PDF...';
        this.disabled = true;
        try {
          const wrap = document.getElementById(`formSection-${activeFormId}`);
          const i18n = (window.APPLICATION_I18N && window.APPLICATION_I18N[currentLang]) 
                    || (window.APPLICATION_I18N && window.APPLICATION_I18N['zh-CN']) 
                    || {};
          await window.ApplicationExporter.exportToPdf(activeFormId, wrap, i18n);
          showToast(i18n.toastPdfExported || '已导出 A4 PDF！');
        } catch (e) {
          console.error(e);
          alert('导出 PDF 失败: ' + e.message);
        } finally {
          this.innerHTML = text;
          this.disabled = false;
        }
      };
    }

    // 原生打印
    const btnPrint = document.getElementById('btnPrint');
    if (btnPrint) {
      btnPrint.onclick = () => window.ApplicationExporter && window.ApplicationExporter.triggerNativePrint();
    }
  }

  // 7. 全局暴露核心方法
  window.switchFormView = switchFormView;
  window.showDirectoryView = showDirectoryView;
  window.updateLanguage = updateLanguage;

  // 8. 双阶段安全初始化 (杜绝 DOMContentLoaded 错过问题)
  function initApp() {
    initSignatureCanvas();
    bindEvents();

    const params = new URLSearchParams(window.location.search);
    const target = params.get('form');
    if (target && formMeta[target]) {
      switchFormView(target);
    } else {
      showDirectoryView();
    }

    // 自动恢复语言设置
    let targetLang = 'zh-CN';
    try {
      const urlLang = params.get('lang');
      const savedLang = localStorage.getItem('spark_form_lang');
      if (urlLang) {
        targetLang = urlLang;
      } else if (savedLang) {
        targetLang = savedLang;
      }
    } catch (e) {}

    if (targetLang && targetLang !== 'zh-CN') {
      updateLanguage(targetLang);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
  } else {
    initApp();
  }

})();

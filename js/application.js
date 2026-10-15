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
    const todayZh = `${new Date().getFullYear()}年${new Date().getMonth()+1}月${new Date().getDate()}日`;
    ['off_date', 'lec_date', 'act_date', 'act_sign_date', 'sec7_date'].forEach(id => {
      const el = document.getElementById(id);
      if (el && !el.value) el.value = el.type === 'date' ? today : todayZh;
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
  function loadDemoData() {
    if (activeFormId === 'office-subsidy') {
      const d = {
        off_date: '2026年10月15日',
        off_lead_name: '陈志远 (David Tan)',
        off_lead_contact: '+65 9182 3746 / @david_spark',
        off_lead_city: '新加坡 (Singapore)',
        off_lead_team: '亚太巅峰团队 (Apex APAC)',
        off_address: '新加坡滨海湾金融中心（MBFC Tower 2）18层 1806室',
        off_area: '150',
        off_capacity: '35',
        off_rent: '4800',
        off_rent_curr: 'USD',
        off_lease_period: '2026年11月01日 至 2027年10月31日',
        off_start_date: '2026-11-01',
        off_sub_amt: '2000',
        off_sub_curr: 'USD',
        off_sub_period: '2026年11月01日 至 2027年04月30日',
        off_comm_name: 'Spark One 亚太核心财富公会 (Telegram官方群)',
        off_comm_link: 'https://t.me/Spark_Apex_APAC',
        off_comm_admins: '陈志远、林敏怡（专属客服）',
        off_comm_members: '520',
        off_comm_active: '日均在线发言 120+ 人，每周研讨参与 350+ 人次',
        off_comm_lang: '新加坡 / 马来西亚 / 常用英语、中文普通话',
        off_comm_daily: '09:00 全球行情早报；14:00 黄金套利盘面解析；21:00 实盘盈利通报',
        off_comm_weekly: '周二晚上线上腾讯会议量化拆解；周六下午线下办公室VIP沙龙交流',
        off_comm_feedback: '建立1对1专属投顾沟通专线，30分钟内快速响应并闭环处理',
        off_kpi_period: '2026年11月01日 至 2026年11月30日',
        off_kpi_goal: '当月直接推荐及团队量化托管总额度突破 250,000 美元'
      };
      Object.keys(d).forEach(k => {
        const el = document.getElementById(k);
        if (el) el.value = d[k];
      });
    } else if (activeFormId === 'lecturer-subsidy') {
      const d = {
        lec_date: '2026年10月15日',
        lec_name: '张峰 (Raymond Zhang)',
        lec_account: 'SPK-882910 / 资深量化主讲导师',
        lec_city: '中国深圳 / 香港',
        lec_contact: '+86 138 2888 9999 / raymond@spark1.cc',
        lec_lang: '中文普通话、粤语、商务英语',
        lec_team: '大湾区国际服务中心',
        lec_ref: '陈志远 (+65 91823746)',
        lec_exp: '8年对冲基金与加密量化培训经验，主讲过超50场高净值投资者财富峰会',
        lec_amt: '2,500 USD / 月',
        lec_audience: '高净值投资人与团队领导人，侧重 MSB 合规存证与量化稳健收益逻辑讲解',
        lec_kpi_period: '2026年11月01日 至 2026年11月30日',
        lec_goal: '全月累计授课时长不少于 18 课时，协助团队达成 30 万美元量化扩容',
        lec_cond: '学员好评率 ≥ 95%，每日按时报备实录视频',
        lec_src: '后台实盘工作台与独立存证系统对账',
        lec_sign: '张峰',
        lec_ref_sign: '陈志远'
      };
      Object.keys(d).forEach(k => {
        const el = document.getElementById(k);
        if (el) el.value = d[k];
      });
    } else if (activeFormId === 'event-apply') {
      const d = {
        act_name: '2026 Spark One 全球多资产配置与AI量化私享会（新加坡站）',
        act_date: '2026-11-20',
        act_venue: '新加坡滨海湾金沙会议展览中心（Sands Expo）4层大宴会厅',
        act_app_name: '李立明 (Marcus Li)',
        act_app_phone: '+65 8299 1188',
        act_app_email: 'marcus.li@sparkunioncapital.com',
        act_app_team: '新加坡狮城创富合伙团队',
        act_theme: '智领财富新纪元：AI黄金量化与全球多资产稳健对冲策略',
        act_b1: '1200',
        act_b2: '600',
        act_b3: '400',
        act_btotal: '2200',
        act_resp1: '李立明 +65 82991188 市场总监',
        act_resp2: '张峰 +86 13828889999 首席导师',
        act_sign: '李立明',
        act_sign_date: '2026-10-15',
        act_p_num: '48',
        act_reg_num: '22',
        act_inv_num: '15',
        act_pic_num: '85',
        act_summary: '全场到会48位高净值客户，现场达成意向资金超38万美元，圆满达成招募考核指标。'
      };
      Object.keys(d).forEach(k => {
        const el = document.getElementById(k);
        if (el) el.value = d[k];
      });
    } else if (activeFormId === 'pilot-grant') {
      const d = {
        sec1_name: '陈志远',
        sec1_account: 'SPK-883921',
        sec1_contact: '+65 9182 3746',
        sec1_residence: '新加坡滨海湾',
        sec1_lang: '中文普通话、英语',
        sec1_team: '亚太巅峰团队',
        sec1_job: '10年量化对冲资产配置背景',
        sec1_hours: '35',
        sec1_perf_detail: '直推 280,000 美元',
        sec2_regions: '新加坡、吉隆坡、中国深圳',
        sec2_reason: '开拓家族办公室，30天内保障达成10倍扶持额度业绩转化',
        sec4_target_clients: '200人',
        sec4_target_events: '4场',
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
          await window.ApplicationExporter.exportToPdf(activeFormId, wrap);
          showToast('已导出 A4 PDF！');
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

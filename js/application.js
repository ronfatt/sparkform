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
    'pilot-grant': '市场领航扶持账号额度申请表',
    'points-mall': '积分商城产品提交与合作申请表'
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
        'pilot-grant': 'pilot_title',
        'points-mall': 'mall_title'
      };
      const key = formKeyMap[formId];
      activeFormTitle.textContent = (key && dict[key]) ? dict[key] : (formMeta[formId] || formId);
    }

    // 默认给各表单日期赋今天的值
    const today = new Date().toISOString().split('T')[0];
    const isZh = (currentLang === 'zh-CN' || currentLang === 'zh-TW');
    const todayStr = isZh ? `${new Date().getFullYear()}年${new Date().getMonth()+1}月${new Date().getDate()}日` : today;
    ['off_date', 'lec_date', 'act_date', 'act_sign_date', 'sec7_date', 'mall_date', 'mall_sign_date'].forEach(id => {
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
    ['off_date', 'lec_date', 'act_date', 'mall_date'].forEach(id => {
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
        'pilot-grant': 'pilot_title',
        'points-mall': 'mall_title'
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
    } else if (activeFormId === 'points-mall') {
      const today = new Date().toISOString().split('T')[0];
      const d = {
        mall_member_name: isZh ? 'Alexander Vance (周启明)' : 'Alexander Vance',
        mall_member_uid: 'SPK-992081',
        mall_country: isZh ? '新加坡 (Singapore)' : 'Singapore',
        mall_company_name: 'Apex Nexus Global Tech Pte. Ltd.',
        mall_reg_city: 'Singapore (UEN: 202419827K)',
        mall_main_category: isZh ? '智能硬件 / Web3冷钱包 / 数码周边' : 'Smart Hardware / Web3 Cold Wallet',
        mall_contact_person: 'Alexander Vance',
        mall_contact_phone: '+65 8299 6688',
        mall_contact_email: 'alex.vance@apexnexus.io',
        mall_contact_im: 'Telegram: @alexvance_sg / WeChat: alex_spark',
        mall_prod_name: isZh ? 'Apex Shield Pro 军工级生物识别多链硬件冷钱包' : 'Apex Shield Pro Biometric Multi-Chain Hardware Cold Wallet',
        mall_brand_name: 'ApexShield',
        mall_origin_cat: 'Singapore / Web3 智能硬件',
        mall_prod_model: 'AS-PRO-2026 / 航空钛铝合金 / CC EAL6+芯片',
        mall_spec_detail: '102×60×7.5mm / 68g / 3.2寸触摸彩屏',
        mall_prod_variants: '深空灰 (Titanium Gray) / 星曜黑 (Obsidian Black)',
        mall_shelf_life: isZh ? '不适用 (原厂硬件三年质保)' : 'N/A (3-Year Factory Warranty)',
        mall_prod_usage: isZh ? '冷端私钥离线签名、助记词多重物理加密、支持 50+ 公链资产安全存储' : 'Offline private key signing, biometric encryption, supports 50+ blockchains',
        mall_pkg_list: isZh ? '冷钱包主机×1、防磁保护套×1、Type-C编织线×1、金属密令卡×3' : 'Host unit x1, Anti-magnetic case x1, Type-C cable x1, Metal mnemonic card x3',
        mall_prod_qty: isZh ? '首批 500 台' : '500 units',
        mall_prod_intro: isZh ? '专为高净值加密资产投资者打造的旗舰级冷钱包，获国际 CC EAL6+ 金融级认证，完全物理断网离线签名，杜绝黑客攻击与钓鱼盗币风险。' : 'Flagship cold wallet for high-net-worth crypto investors, featuring CC EAL6+ financial certification and 100% air-gapped offline signing.',
        mall_cost_prod: '85',
        mall_cost_pack: '3',
        mall_cost_ship: '12',
        mall_retail_price: '$168.00 USD',
        mall_supply_price: '$95.00 USD',
        mall_price_valid: '2026-12-31',
        mall_stock_qty: isZh ? '500台 (现货稳定，月产能 3,000 台)' : '500 units (Monthly cap 3,000)',
        mall_restock_lead: isZh ? '起订量: 1台 / 补货: 7个工作日' : 'MOQ: 1 unit / Lead time: 7 days',
        mall_ship_origin: isZh ? '新加坡自由港仓库 / 香港国际物流中心' : 'Singapore FreePort / Hong Kong Logistics Hub',
        mall_ship_dest: isZh ? '新加坡、马来西亚、中国香港/台湾、印尼、泰国、日本、韩国、阿联酋、欧洲主要国家' : 'Singapore, Malaysia, HK/TW, Indonesia, Thailand, Japan, Korea, UAE, EU',
        mall_ship_restricted: isZh ? '极偏远未通邮岛屿及受海关禁运国家' : 'Sanctioned territories and remote unserviced islands',
        mall_ship_time: isZh ? '下单后24小时内闪电发货，亚太地区3-5天送达，欧美7-10天' : 'Dispatched within 24h, APAC 3-5 days, EU/US 7-10 days',
        mall_customs_payer: isZh ? '合作方包税包清关 (DDP完税交付)' : 'DDP (Partner covers all customs & taxes)',
        mall_return_policy: isZh ? '支持7天无理由退换 (未激活塑封完好)；3年内非人为故障免费换新' : '7-day unconditional return (unopened); 3-year warranty replacement',
        mall_damage_handling: isZh ? '运输途中破损/丢失全额极速补发，由我司承担' : 'Instant free reshipment for damaged/lost packages',
        mall_aftersales_contact: '+65 8299 6688 / support@apexnexus.io (4h内响应)',
        mall_cert_files: 'CE, FCC, RoHS, CC EAL6+, 商标注册证书, 品牌授权书',
        mall_sign: 'Alexander Vance',
        mall_sign_date: today
      };
      Object.keys(d).forEach(k => {
        const el = document.getElementById(k);
        if (el) el.value = d[k];
      });
      updatePointsCalc();
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

  // 积分商城实时联动计算
  function updatePointsCalc() {
    const prod = parseFloat(document.getElementById('mall_cost_prod')?.value) || 0;
    const pack = parseFloat(document.getElementById('mall_cost_pack')?.value) || 0;
    const ship = parseFloat(document.getElementById('mall_cost_ship')?.value) || 0;
    const total = prod + pack + ship;
    const points = Math.round(total * 800);
    const totalEl = document.getElementById('mall_cost_total');
    const pointsEl = document.getElementById('mall_calc_points');
    const apprEl = document.getElementById('mall_appr_points');
    if (totalEl) totalEl.value = total > 0 ? `$${total.toFixed(2)} USD` : '';
    if (pointsEl) pointsEl.value = points > 0 ? `${points.toLocaleString()} 积分` : '';
    if (apprEl && points > 0) {
      apprEl.value = `${points.toLocaleString()} 积分`;
    }
  }

  // 7. 全局暴露核心方法
  window.switchFormView = switchFormView;
  window.showDirectoryView = showDirectoryView;
  window.updateLanguage = updateLanguage;
  window.updatePointsCalc = updatePointsCalc;

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

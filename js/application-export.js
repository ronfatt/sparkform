/**
 * Spark One 申请表导出引擎 (Word & PDF)
 * 严格按照原版 PDF 规格排版：
 * - 办公室补贴申请表: 严格 2 页 (Page 1 + Page 2)
 * - 讲师补贴申请表: 严格 2 页 (Page 1 + Page 2)
 * - 活动申请表: 严格 1 页 (Single Page)
 * - 市场领航扶持申请表: 紧凑 A4
 * 全面支持 12 国语言本地化导出 (中文/繁体/英文/韩文/日文/泰文/越文/印尼文/马来文/西文/葡文/阿文)
 */

window.ApplicationExporter = (function () {
  'use strict';

  function getTodayString(lang) {
    const d = new Date();
    const isZh = !lang || lang === 'zh-CN' || lang === 'zh-TW';
    if (isZh) {
      return `${d.getFullYear()}年${String(d.getMonth() + 1).padStart(2, '0')}月${String(d.getDate()).padStart(2, '0')}日`;
    }
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function getFileName(formId, ext, t) {
    const d = new Date().toISOString().split('T')[0];
    const translate = t || ((k, fallback) => fallback);
    let title = "Spark";
    if (formId === 'office-subsidy') title = translate('off_title', '办公室补贴申请表');
    else if (formId === 'lecturer-subsidy') title = translate('lec_title', '讲师补贴申请表');
    else if (formId === 'event-apply') title = translate('act_title', '活动申请表');
    else title = translate('pilot_title', '市场领航扶持账号申请表');
    
    const cleanTitle = title.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '_');
    return `Spark_${cleanTitle}_${d}.${ext}`;
  }

  // =========================================================================
  // 1. 导出 WORD (.docx)
  // =========================================================================
  async function exportToWord(formId, formData, i18nDict) {
    const dict = i18nDict || (window.APPLICATION_I18N && window.APPLICATION_I18N[window.currentAppLang || 'zh-CN']) || {};
    const t = (k, fallback) => (dict[k] !== undefined ? dict[k] : (fallback !== undefined ? fallback : k));
    const currentLang = window.currentAppLang || 'zh-CN';

    if (!window.docx) {
      alert(t('statusLoadingWord', 'Word 导出组件加载中，请稍后再试...'));
      return;
    }

    const {
      Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
      WidthType, BorderStyle, AlignmentType, ShadingType, PageBreak
    } = window.docx;

    // 统一样式常数
    const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: "94A3B8" };
    const tableBorders = {
      top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder,
      insideHorizontal: thinBorder, insideVertical: thinBorder
    };

    const purpleHeader = { fill: "2E1065", type: ShadingType.CLEAR };
    const grayBg = { fill: "F8FAFC", type: ShadingType.CLEAR };

    // 创建标准双列行 (项目 | 填写内容)
    function createTwoColRow(label, value, labelWidth = 32) {
      return new TableRow({
        children: [
          new TableCell({
            width: { size: labelWidth, type: WidthType.PERCENTAGE },
            shading: grayBg,
            margins: { top: 50, bottom: 50, left: 100, right: 100 },
            children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 18, font: "Microsoft YaHei" })] })]
          }),
          new TableCell({
            width: { size: 100 - labelWidth, type: WidthType.PERCENTAGE },
            margins: { top: 50, bottom: 50, left: 100, right: 100 },
            children: [new Paragraph({ children: [new TextRun({ text: value || '', size: 18, font: "Microsoft YaHei" })] })]
          })
        ]
      });
    }

    // 创建章节紫底标头
    function createSectionBar(title) {
      return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: tableBorders,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                shading: purpleHeader,
                margins: { top: 60, bottom: 60, left: 100, right: 100 },
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: title, bold: true, color: "FFFFFF", size: 19, font: "Microsoft YaHei" })]
                  })
                ]
              })
            ]
          })
        ]
      });
    }

    // 创建表头 (Logo 文字 + 大标题 + 英文副标题 + 日期类型)
    function createDocHeader(mainTitle, subTitle, dateStr, typeStr) {
      return [
        new Paragraph({
          children: [
            new TextRun({ text: "SPARK UNION CAPITAL INC.", bold: true, size: 20, color: "2E1065", font: "Space Grotesk" }),
            new TextRun({ text: "                                                                           ///", color: "D97706", bold: true, size: 20 })
          ],
          spacing: { after: 60 }
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: mainTitle, bold: true, size: 36, font: "Microsoft YaHei", color: "1E1B4B" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: subTitle, size: 15, bold: true, color: "64748B", font: "Space Grotesk" })],
          spacing: { after: 100 }
        }),
        new Paragraph({
          children: [
            new TextRun({ text: `${t('app_date_lbl', '申请日期：')} ${dateStr || getTodayString(currentLang)}        `, size: 17, font: "Microsoft YaHei" }),
            new TextRun({ text: `${t('app_type_lbl', '申请类型：')} ${typeStr || ('☑ ' + t('type_first', '首次申请') + '   ☐ ' + t('type_renew', '续期申请') + '   ☐ ' + t('type_change', '信息变更'))}`, size: 17, font: "Microsoft YaHei" })
          ],
          spacing: { after: 80 }
        })
      ];
    }

    let docChildren = [];

    // =========================================================================
    // 表单 A: 办公室补贴申请表 (严格 2 页)
    // =========================================================================
    if (formId === 'office-subsidy') {
      const tFirst = t('type_first', '首次申请');
      const tRenew = t('type_renew', '续期申请');
      const tChange = t('type_change', '信息变更');
      const typeText = formData.off_type === 'renew' ? `☐ ${tFirst}   ☑ ${tRenew}   ☐ ${tChange}` : (formData.off_type === 'change' ? `☐ ${tFirst}   ☐ ${tRenew}   ☑ ${tChange}` : `☑ ${tFirst}   ☐ ${tRenew}   ☐ ${tChange}`);

      const optDone = t('opt_done', '已完成');
      const optDoing = t('opt_doing', '筹备中');
      const optNone = t('opt_none', '未准备');

      // --- PAGE 1 ---
      docChildren.push(
        ...createDocHeader(t('off_title', "办公室补贴申请表"), t('off_sub', "OFFICE SUBSIDY APPLICATION"), formData.off_date, typeText),
        createSectionBar(t('off_sec1', "一、负责人及办公室信息")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('off_f_lead', "负责人姓名"), formData.off_lead_name),
            createTwoColRow(t('off_f_contact', "联系电话∕社交账号"), formData.off_lead_contact),
            createTwoColRow(t('off_f_city', "所属国家∕城市"), formData.off_lead_city),
            createTwoColRow(t('off_f_team', "团队名称"), formData.off_lead_team),
            createTwoColRow(t('off_f_addr', "办公室详细地址"), formData.off_address),
            createTwoColRow(t('off_f_area', "办公室面积"), `${formData.off_area || '______'} ${t('unit_sqm', '平方米')}`),
            createTwoColRow(t('off_f_cap', "可容纳人数"), `${formData.off_capacity || '______'} ${t('unit_person', '人')}`),
            createTwoColRow(t('off_f_rent', "每月租金"), `${t('off_lbl_amt', '金额')} ${formData.off_rent || '______'}，${t('unit_usd', '币种：')}${formData.off_rent_curr || 'USD'}`),
            createTwoColRow(t('off_f_lease', "租赁期限"), formData.off_lease_period || "______"),
            createTwoColRow(t('off_f_start', "正式启用日期"), formData.off_start_date || "______"),
            createTwoColRow(t('off_f_sub', "申请补贴"), `${t('unit_month', '每月')} ${formData.off_sub_amt || '______'}，${t('unit_usd', '币种：')}${formData.off_sub_curr || 'USD'}`),
            createTwoColRow(t('off_f_sub_p', "申请补贴期限"), formData.off_sub_period || "______")
          ]
        }),
        new Paragraph({ spacing: { before: 100, after: 40 } }),
        createSectionBar(t('off_sec2', "二、物料准备情况")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: purpleHeader, margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_mat', "物料项目"), bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 35, type: WidthType.PERCENTAGE }, shading: purpleHeader, margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_stat', "准备情况"), bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, shading: purpleHeader, margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_qty', "数量∕说明"), bold: true, color: "FFFFFF", size: 17 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_m1', "公司标识、门牌及背景墙"), size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☑${optDone}  ☐${optDoing}  ☐${optNone}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_1 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_m2', "公司介绍及Spark One介绍资料"), size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☑${optDone}  ☐${optDoing}  ☐${optNone}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_2 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_m3', "经审核的课程课件及讲解资料"), size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☑${optDone}  ☐${optDoing}  ☐${optNone}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_3 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_m4', "风险说明及客户服务指引"), size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☑${optDone}  ☐${optDoing}  ☐${optNone}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_4 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_m5', "桌椅、投影∕显示设备"), size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☑${optDone}  ☐${optDoing}  ☐${optNone}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_5 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_m6', "网络、收音及视频拍摄设备"), size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☑${optDone}  ☐${optDoing}  ☐${optNone}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_6 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_m7', "签到表、活动记录及咨询登记表"), size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☑${optDone}  ☐${optDoing}  ☐${optNone}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_7 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_m8', "其他物料"), size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☐${optDone}  ☑${optDoing}  ☐${optNone}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_8 || "", size: 16 })] })] })] })
          ]
        }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "01 / 02", size: 16, bold: true, color: "64748B" })], spacing: { before: 120 } }),
        
        // --- 强制分页: 产生第 2 页 ---
        new Paragraph({ children: [new PageBreak()] }),

        // --- PAGE 2 ---
        new Paragraph({
          children: [
            new TextRun({ text: "SPARK UNION CAPITAL INC.", bold: true, size: 20, color: "2E1065" }),
            new TextRun({ text: `                     ${t('off_title', '办公室补贴申请表')}`, bold: true, size: 24, color: "1E1B4B" })
          ],
          spacing: { after: 60 }
        }),
        createSectionBar(t('off_sec3', "三、社区群运营情况")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('off_c_name', "社区群名称及平台"), formData.off_comm_name),
            createTwoColRow(t('off_c_link', "社区群链接∕群编号"), formData.off_comm_link),
            createTwoColRow(t('off_c_admin', "群主及管理员"), formData.off_comm_admins),
            createTwoColRow(t('off_c_members', "当前群人数"), formData.off_comm_members),
            createTwoColRow(t('off_c_active', "近30天活跃人数及统计口径"), formData.off_comm_active),
            createTwoColRow(t('off_c_lang', "服务地区及主要语言"), formData.off_comm_lang),
            createTwoColRow(t('off_c_daily', "每日内容发布安排"), formData.off_comm_daily),
            createTwoColRow(t('off_c_weekly', "每周课程及活动安排"), formData.off_comm_weekly),
            createTwoColRow(t('off_c_feed', "客户问题反馈与处理方式"), formData.off_comm_feedback)
          ]
        }),
        new Paragraph({ spacing: { before: 80, after: 40 } }),
        createSectionBar(t('off_sec4', "四、每日讲课视频提交要求")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Paragraph({ children: [new TextRun({ text: t('off_r1', "1. 提交频次：每个批准的运营日提交不少于 1 条当日实际授课视频；休息日及停课情况提前报备。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('off_r2', "2. 视频时长：每条不少于 15 分钟；完整课程或课堂片段＋完整录像留档。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('off_r3', "3. 视频内容：注明日期、办公室名称、讲师姓名、课程主题及参与人数，清晰展示实际授课情况。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('off_r4', "4. 课程范围：公司介绍、产品规则、平台操作、风险教育及客户服务；使用经审核的资料。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('off_r5', "5. 提交方式：每日 21:00 前提交至官方审核群，同步附上签到人数及课程简报。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('off_r6', "6. 真实性要求：不得使用旧视频重复提交、拼接伪造授课记录；拍摄参与者前取得同意。"), size: 16 })], spacing: { after: 80 } }),
        createSectionBar(t('off_sec5', "五、考核期限与目标")),
        new Paragraph({
          children: [
            new TextRun({ text: `${t('kpi_period_lbl', '考核期限：')}${formData.off_kpi_period || '______'}        `, size: 16 }),
            new TextRun({ text: `${t('kpi_freq_lbl', '考核频率：')}☐${t('freq_weekly', '每周')}  ☑${t('freq_monthly', '每月')}  ☐${t('freq_other', '其他')}`, size: 16 })
          ],
          spacing: { before: 40, after: 30 }
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 34, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_kpi', "考核指标"), bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_target', "批准目标"), bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_actual', "实际完成"), bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_res', "审核结果"), bold: true, color: "FFFFFF", size: 17 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_k1', "办公室实际开放天数"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `≥ 24 ${t('unit_day', '天')}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_k2', "课程∕活动场次"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `≥ 8 ${t('unit_session', '场')}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_k3', "有效视频提交天数"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `≥ 20 ${t('unit_day', '天')}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_k4', "课程参与人数∕人次"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `≥ 150 ${t('unit_person_time', '人次')}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_k5', "社区活跃人数"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `≥ 100 ${t('unit_person', '人')}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_k6', "客户咨询及问题处理完成率"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 98%", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('off_k7', "业绩目标："), bold: true, size: 16 })] })] }), new TableCell({ columnSpan: 3, children: [new Paragraph({ children: [new TextRun({ text: formData.off_kpi_goal || "", size: 16 })] })] })] })
          ]
        }),
        new Paragraph({
          children: [new TextRun({ text: `${t('off_att_title', '附件清单：')} ☑${t('off_a1', '租赁合同')}  ☑${t('off_a2', '租金支付凭证')}  ☑${t('off_a3', '办公室现场照片')}  ☑${t('off_a4', '物料照片')}  ☑${t('off_a5', '社区群信息截图')}  ☑${t('off_a6', '课程计划')}  ☐${t('off_a7', '其他')}`, size: 15, color: "334155" })],
          spacing: { before: 50, after: 60 }
        }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "02 / 02", size: 16, bold: true, color: "64748B" })] })
      );

    // =========================================================================
    // 表单 B: 讲师补贴申请表 (严格 2 页)
    // =========================================================================
    } else if (formId === 'lecturer-subsidy') {
      const tLecFirst = t('type_first', '首次申请');
      const tLecRenew = t('type_renew', '续期申请');
      const tLecChange = t('type_change', '信息变更');
      const typeText = formData.lec_type === 'renew' ? `☐ ${tLecFirst}   ☑ ${tLecRenew}   ☐ ${tLecChange}` : (formData.lec_type === 'change' ? `☐ ${tLecFirst}   ☐ ${tLecRenew}   ☑ ${tLecChange}` : `☑ ${tLecFirst}   ☐ ${tLecRenew}   ☐ ${tLecChange}`);

      // --- PAGE 1 ---
      docChildren.push(
        ...createDocHeader(t('lec_title', "讲师补贴申请表"), t('lec_sub', "LECTURER SUBSIDY APPLICATION"), formData.lec_date, typeText),
        createSectionBar(t('lec_sec1', "一、讲师基本信息")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('lec_f_name', "讲师姓名"), formData.lec_name),
            createTwoColRow(t('lec_f_account', "会员编号∕职务"), formData.lec_account),
            createTwoColRow(t('lec_f_city', "所属国家∕城市"), formData.lec_city),
            createTwoColRow(t('lec_f_contact', "联系电话∕社交账号"), formData.lec_contact),
            createTwoColRow(t('lec_f_lang', "授课语言"), formData.lec_lang),
            createTwoColRow(t('lec_f_team', "所属团队∕办公室"), formData.lec_team),
            createTwoColRow(t('lec_f_ref', "推荐负责人及联系方式"), formData.lec_ref),
            createTwoColRow(t('lec_f_exp', "授课经验及擅长主题"), formData.lec_exp),
            createTwoColRow(t('lec_f_amt', "申请补贴金额及币种"), formData.lec_amt),
            createTwoColRow(t('lec_f_mode', "申请补贴方式"), `☑${t('lec_m_half', '按半月')}   ☐${t('lec_m_session', '按场')}   ☐${t('lec_m_other', '其他')}`)
          ]
        }),
        new Paragraph({ spacing: { before: 80, after: 30 } }),
        createSectionBar(t('lec_sec2', "二、服务社区")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_cname', "社区名称"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 18, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_cid', "平台∕群编号"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 16, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_area', "服务地区"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 12, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_num', "群人数"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 16, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_lead', "管理员∕负责人"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 16, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_freq', "计划频次"), bold: true, color: "FFFFFF", size: 16 })] })] })
              ]
            }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c1_name || "Apex APAC Wealth Group", size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c1_id || "TG @spark_apex", size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c1_area || "Singapore / Malaysia", size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${formData.lec_c1_num || '500'} ${t('unit_person', '人')}`, size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c1_lead || "", size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c1_freq || "", size: 16 })] })] })
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c2_name || "Quant Private Group", size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c2_id || "WeChat Group 08", size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c2_area || "Shenzhen / HK", size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${formData.lec_c2_num || '380'} ${t('unit_person', '人')}`, size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c2_lead || "", size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.lec_c2_freq || "", size: 16 })] })] })
            ] })
          ]
        }),
        new Paragraph({ children: [new TextRun({ text: `${t('lec_aud_lbl', '主要服务对象及需求：')}${formData.lec_audience || ""}`, size: 16 })], spacing: { before: 40, after: 60 } }),
        createSectionBar(t('lec_sec3', "三、授课计划")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_dir', "课程方向"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_theme', "拟讲主题"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 15, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_cnt', "计划场次"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 15, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_dur', "单场时长"), bold: true, color: "FFFFFF", size: 16 })] })] })
              ]
            }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_d1', "公司介绍与发展规划"), size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_theme_1', "企业愿景与MSB合规架构"), size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `4 ${t('unit_session', '场')}`, size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `45 ${t('unit_min', '分钟')}`, size: 16 })] })] })
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_d2', "Spark One 产品及规则说明"), size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_theme_2', "黄金与加密量化模型及稳健收益拆解"), size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `6 ${t('unit_session', '场')}`, size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `60 ${t('unit_min', '分钟')}`, size: 16 })] })] })
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_d3', "平台操作与客户服务"), size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_theme_3', "DApp工作台充提与质押实操"), size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `4 ${t('unit_session', '场')}`, size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `30 ${t('unit_min', '分钟')}`, size: 16 })] })] })
            ] }),
            new TableRow({ children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_d4', "投资风险及账户安全教育"), size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_theme_4', "私钥保管、防钓鱼诈骗及理性投资指引"), size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `2 ${t('unit_session', '场')}`, size: 16 })] })] }), 
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `40 ${t('unit_min', '分钟')}`, size: 16 })] })] })
            ] })
          ]
        }),
        new Paragraph({
          children: [
            new TextRun({ text: `${t('lec_ver_lbl', '课件版本 / 审核负责人：')} V5.2 / ${t('appr_team', '官方审核组')}        `, size: 16 }),
            new TextRun({ text: `${t('lec_form_lbl', '授课形式：')}☑${t('lec_f_offline', '线下')}  ☑${t('lec_f_live', '线上直播')}  ☐${t('lec_f_record', '录播')}  ☐${t('lec_f_hybrid', '混合')}`, size: 16 })
          ],
          spacing: { before: 40, after: 80 }
        }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "01 / 02", size: 16, bold: true, color: "64748B" })] }),

        // --- 强制分页: 产生第 2 页 ---
        new Paragraph({ children: [new PageBreak()] }),

        // --- PAGE 2 ---
        new Paragraph({
          children: [
            new TextRun({ text: "SPARK UNION CAPITAL INC.", bold: true, size: 20, color: "2E1065" }),
            new TextRun({ text: `                     ${t('lec_title', '讲师补贴申请表')}`, bold: true, size: 24, color: "1E1B4B" })
          ],
          spacing: { after: 60 }
        }),
        createSectionBar(t('lec_sec4', "四、每日视频规则")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Paragraph({ children: [new TextRun({ text: t('lec_r1', "1. 提交频次：每个批准的授课日提交不少于 1 条本人当日授课视频。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('lec_r2', "2. 视频时长：每条不少于 15 分钟；完整课程或课程片段+完整录像留档。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('lec_r3', "3. 视频标注：讲师姓名、授课日期、服务社区、课程主题及实际参与人数。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('lec_r4', "4. 质量要求：画面与声音清晰，内容完整连贯，能够核验实际授课情况。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('lec_r5', "5. 提交时间：每日 22:00 前提交至官方讲师考核专群，附课程要点、参与人数及问题反馈。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('lec_r6', "6. 内容要求：使用经审核的课件，准确说明产品规则及风险；不得承诺保本收益。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('lec_r7', "7. 记录要求：不得以旧视频冒充当日课程；同一场课程服务多个社区时须注明。"), size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: t('lec_r8', "8. 隐私要求：公开使用含学员面部、姓名或联系方式的视频前，取得相应同意。"), size: 16 })], spacing: { after: 60 } }),
        createSectionBar(t('lec_sec5', "五、考核期限与业绩目标")),
        new Paragraph({
          children: [
            new TextRun({ text: `${t('kpi_period_lbl', '考核期限：')}${formData.lec_kpi_period || '______'}        `, size: 16 }),
            new TextRun({ text: `${t('kpi_freq_lbl', '考核频率：')}☐${t('freq_weekly', '每周')}  ☑${t('freq_monthly', '每月')}  ☐${t('freq_other', '其他')}`, size: 16 })
          ],
          spacing: { before: 30, after: 30 }
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 34, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_kpi', "考核指标"), bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_target', "批准目标"), bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_actual', "实际完成"), bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('off_th_res', "审核结果"), bold: true, color: "FFFFFF", size: 17 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_k1', "实际授课天数"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `≥ 20 ${t('unit_day', '天')}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_k2', "有效授课场次"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `≥ 16 ${t('unit_session', '场')}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_k3', "有效视频提交率"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "100%", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_k4', "服务社区数量"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 2", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_k5', "课程参与人数∕人次"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `≥ 300 ${t('unit_person_time', '人次')}`, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_k6', "学员满意度"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 95%", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('lec_k7', "课件准确性及讲解质量"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('kpi_qualified', "合格"), size: 16 })] })] })] })
          ]
        }),
        new Paragraph({ spacing: { before: 60, after: 30 } }),
        createSectionBar(t('lec_sec6', "六、附件及申请人确认")),
        new Paragraph({
          children: [new TextRun({ text: `${t('lec_att_title', '附件：')} ☑${t('lec_att_1', '个人简介')}  ☑${t('lec_att_2', '试讲视频')}  ☑${t('lec_att_3', '课程计划')}  ☑${t('lec_att_4', '课件样本')}  ☑${t('lec_att_5', '服务社区证明')}  ☐${t('lec_att_6', '其他')}`, size: 15 })],
          spacing: { before: 30, after: 30 }
        }),
        new Paragraph({
          children: [new TextRun({ text: t('lec_confirm_txt', "申请人确认：本人确认申请资料真实，按批准计划授课并提交记录，遵守课件审核及宣传要求。无法按计划授课时，提前向负责人报备。"), size: 15, color: "334155" })],
          spacing: { after: 50 }
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_sign', "申请人签名"), bold: true, size: 16 })] })] }),
                new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: formData.lec_sign || "", size: 16 })] })] }),
                new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_ref', "推荐负责人"), bold: true, size: 16 })] })] }),
                new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: formData.lec_ref_sign || "", size: 16 })] })] })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({ shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_appr', "审批结果"), bold: true, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `☑${t('opt_pass', '批准')}  ☐${t('opt_modify', '需修改')}  ☐${t('opt_reject', '不批准')}`, size: 15 })] })] }),
                new TableCell({ shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_aud', "审核负责人"), bold: true, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('appr_team', "官方审核组"), size: 16 })] })] })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({ shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_date', "审核日期"), bold: true, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: getTodayString(currentLang), size: 16 })] })] }),
                new TableCell({ shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: t('lec_th_rem', "备注"), bold: true, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('appr_pass_full', '予以全额批准'), size: 16 })] })] })
              ]
            })
          ]
        }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "02 / 02", size: 16, bold: true, color: "64748B" })], spacing: { before: 80 } })
      );

    // =========================================================================
    // 表单 C: SPARK UNION CAPITAL 活动申请表 (严格 1 页)
    // =========================================================================
    } else if (formId === 'event-apply') {
      docChildren.push(
        new Paragraph({
          children: [
            new TextRun({ text: "SPARK UNION CAPITAL INC.", bold: true, size: 20, color: "2E1065" }),
            new TextRun({ text: `                     ${t('act_main_title', '活动申请表')}`, bold: true, size: 26, color: "1E1B4B" })
          ],
          spacing: { after: 60 }
        }),
        createSectionBar(t('act_sec1', "1. 活动基本信息")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('act_f_name', "活动名称"), formData.act_name, 22),
            createTwoColRow(t('act_f_type', "活动类型"), `☑${t('act_t1', '会员交流会')}  ☑${t('act_t2', '财富分享会')}  ☐${t('act_t3', '新人说明会')}  ☐${t('act_t4', '市场推广活动')}  ☐${t('act_t5', '社区沙龙')}  ☐${t('act_t6', '公益活动')}`, 22),
            createTwoColRow(t('act_f_date', "活动日期"), formData.act_date || getTodayString(currentLang), 22),
            createTwoColRow(t('act_f_venue', "活动地点"), formData.act_venue, 22),
            createTwoColRow(t('act_f_ppl', "预计参与人数"), `☐${t('act_ppl_1', '<10人')}  ☐${t('act_ppl_2', '10-30人')}  ☑${t('act_ppl_3', '30-50人')}  ☐${t('act_ppl_4', '50-100人')}  ☐${t('act_ppl_5', '100人以上')}`, 22)
          ]
        }),
        new Paragraph({ spacing: { before: 40, after: 20 } }),
        createSectionBar(t('act_sec2', "2. 申请人信息")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('act_f_appname', "申请人姓名"), formData.act_app_name, 22),
            createTwoColRow(t('act_f_phone', "联系电话"), formData.act_app_phone, 22),
            createTwoColRow(t('act_f_email', "电子邮箱"), formData.act_app_email, 22),
            createTwoColRow(t('act_f_team', "所属市场团队"), formData.act_app_team, 22)
          ]
        }),
        new Paragraph({ spacing: { before: 40, after: 20 } }),
        createSectionBar(t('act_sec3', "3. 活动内容说明")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('act_f_theme', "活动主题"), formData.act_theme, 22),
            createTwoColRow(t('act_f_purpose', "活动目的"), `☑${t('act_p1', '品牌宣传')}  ☑${t('act_p2', '新会员招募')}  ☑${t('act_p3', '市场拓展')}  ☑${t('act_p4', '产品介绍')}  ☐${t('act_p5', '投资知识分享')}  ☐${t('act_p6', '团队建设')}`, 22),
            createTwoColRow(t('act_f_flow', "活动流程简介"), formData.act_agenda || `1. ${t('act_flow_1', '嘉宾签到与宣传片暖场播放 (30分钟)')}\n2. ${t('act_flow_2', 'Spark One 核心量化策略与MSB合规架构介绍 (45分钟)')}\n3. ${t('act_flow_3', '实盘工作台演示与现场QA答疑 (30分钟)')}\n4. ${t('act_flow_4', '商务茶歇与意向VIP开单辅导 (45分钟)')}`, 22)
          ]
        }),
        new Paragraph({ spacing: { before: 40, after: 20 } }),
        createSectionBar(t('act_sec4', "4. 活动预算申请")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 60, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('act_th_item', "项目"), bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: t('act_th_budget', "预算金额 (USD)"), bold: true, color: "FFFFFF", size: 16 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('act_b1_lbl', "场地租金及设备费"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `$${formData.act_b1 || '1200'}`, size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('act_b2_lbl', "餐饮茶歇及水果点心"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `$${formData.act_b2 || '600'}`, size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('act_b3_lbl', "宣传物料、易拉宝及伴手礼"), size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `$${formData.act_b3 || '400'}`, size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t('act_b_total', "合计"), bold: true, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `$${formData.act_btotal || '2200'} USD`, bold: true, size: 16, color: "DC2626" })] })] })] })
          ]
        }),
        new Paragraph({
          children: [new TextRun({ text: `${t('act_sec5', '5. 物料支持申请')} ☑${t('act_m1', '企业宣传册')}  ☑${t('act_m2', '品牌易拉宝')}  ☐${t('act_m3', '活动海报')}  ☑${t('act_m4', '宣传视频')}  ☑${t('act_m5', '产品介绍PPT')}  ☐${t('act_m6', '礼品支持')}`, size: 15 })],
          spacing: { before: 30, after: 30 }
        }),
        new Paragraph({
          children: [new TextRun({ text: `${t('act_sec6', '6. 活动负责人')} ${t('act_resp_1', '第一负责人')}：${formData.act_resp1 || ''}      ${t('act_resp_2', '第二负责人')}：${formData.act_resp2 || ''}`, size: 15 })],
          spacing: { after: 30 }
        }),
        new Paragraph({
          children: [new TextRun({ text: `${t('act_sec7', '7. 风险管理承诺')} ${t('act_risk_txt', '活动严格遵守当地法律法规；不进行虚假宣传；不承诺保本固定收益；维护星火联合资本品牌形象。')}`, size: 14, color: "475569" })],
          spacing: { after: 30 }
        }),
        new Paragraph({
          children: [
            new TextRun({ text: `${t('act_sign_lbl', '申请人签字：')} ${formData.act_sign || ''}                     `, bold: true, size: 16 }),
            new TextRun({ text: `${t('act_date_lbl', '日期：')} ${formData.act_sign_date || getTodayString(currentLang)}`, bold: true, size: 16 })
          ],
          spacing: { after: 40 }
        }),
        createSectionBar(t('act_sec8', "8. 活动结案反馈（活动结束后填写）")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${t('act_f_realppl', '实际参与人数：')} ${formData.act_p_num || '0'} ${t('unit_person', '人')}`, size: 15 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${t('act_f_reg', '新增注册会员：')} ${formData.act_reg_num || '0'} ${t('unit_person', '人')}`, size: 15 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${t('act_f_inv', '新增投资会员：')} ${formData.act_inv_num || '0'} ${t('unit_person', '人')}`, size: 15 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${t('act_f_pic', '活动照片数量：')} ${formData.act_pic_num || '0'} ${t('unit_sheet', '张')}`, size: 15 })] })] })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({ columnSpan: 4, children: [new Paragraph({ children: [new TextRun({ text: `${t('act_f_sum', '活动总结：')} ${formData.act_summary || ''}`, size: 15 })] })] })
              ]
            })
          ]
        })
      );

    // =========================================================================
    // 表单 D: 市场领航扶持申请表 (紧凑 A4)
    // =========================================================================
    } else {
      docChildren.push(
        new Paragraph({
          children: [
            new TextRun({ text: "SPARK UNION CAPITAL INC.", bold: true, size: 20, color: "2E1065" }),
            new TextRun({ text: `           ${t('pilot_title', '市场领航扶持账号申请表')}`, bold: true, size: 26, color: "1E1B4B" })
          ],
          spacing: { after: 60 }
        }),
        createSectionBar(t('pilot_sec1', "一、申请人基本资料")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('pilot_name', "姓名"), formData.sec1_name, 25),
            createTwoColRow(t('pilot_account', "会员账号"), formData.sec1_account, 25),
            createTwoColRow(t('pilot_contact', "联系电话/邮箱"), formData.sec1_contact, 25),
            createTwoColRow(t('pilot_residence', "常驻国家/城市"), formData.sec1_residence, 25),
            createTwoColRow(t('pilot_hours', "每周投入时间"), `${formData.sec1_hours || '35'} ${t('unit_hours_week', '小时/周')}`, 25),
            createTwoColRow(t('pilot_perf', "业绩情况"), formData.sec1_perf_detail, 25)
          ]
        }),
        new Paragraph({ spacing: { before: 30, after: 20 } }),
        createSectionBar(t('pilot_sec2', "二、申请扶持账号额度")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('pilot_amount', "申请额度"), "$10,000 USD", 25),
            createTwoColRow(t('pilot_target_city', "目标开拓城市"), formData.sec2_regions, 25),
            createTwoColRow(t('pilot_reason', "申请理由与规划"), formData.sec2_reason, 25)
          ]
        }),
        new Paragraph({ spacing: { before: 30, after: 20 } }),
        createSectionBar(t('pilot_sec4', "三、30天预期目标与声明")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow(t('pilot_reach', "触达客户人数"), formData.sec4_target_clients, 25),
            createTwoColRow(t('pilot_events', "举办说明会场次"), formData.sec4_target_events, 25),
            createTwoColRow(t('pilot_compliance_title', "合规承诺与签署"), `${t('pilot_pledge', '本人确认提交信息真实，未达扶持额度10倍业绩将自觉接受账号收回；各团队长业绩不重复计算。')}\n${t('sign_lbl', '申请人签名：')}${formData.sec1_name || ''}        ${t('date_lbl', '日期：')}${getTodayString(currentLang)}`, 25)
          ]
        })
      );
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 600, right: 600, bottom: 600, left: 600 }
          }
        },
        children: docChildren
      }]
    });

    const blob = await Packer.toBlob(doc);
    const fileName = getFileName(formId, 'docx', t);
    window.saveAs(blob, fileName);
    return fileName;
  }

  // =========================================================================
  // 2. 导出 PDF (.pdf) - 严格按 .a4-paper 逐页 1:1 独立绘制，彻底根除任何多余空白页！
  // =========================================================================
  async function exportToPdf(formId, targetElement, i18nDict) {
    const dict = i18nDict || (window.APPLICATION_I18N && window.APPLICATION_I18N[window.currentAppLang || 'zh-CN']) || {};
    const t = (k, fallback) => (dict[k] !== undefined ? dict[k] : (fallback !== undefined ? fallback : k));
    const fileName = getFileName(formId, 'pdf', t);
    
    // 获取当前表单下的所有真实 A4 页面
    const papers = targetElement.querySelectorAll('.a4-paper');
    if (!papers || papers.length === 0) {
      window.print();
      return;
    }

    const jsPDFConstructor = (window.jspdf && window.jspdf.jsPDF) || window.jsPDF;
    
    if (window.html2canvas && jsPDFConstructor) {
      // 创建精准 A4 PDF 实例 (210mm x 297mm, portrait)
      const pdf = new jsPDFConstructor('p', 'mm', 'a4');
      const pdfWidth = 210;
      const pdfHeight = 297;

      for (let i = 0; i < papers.length; i++) {
        const paper = papers[i];
        
        // 克隆该单张 A4 纸
        const clone = paper.cloneNode(true);

        // 将输入框值直接转化为文本，防止光标或阴影
        clone.querySelectorAll('input, textarea').forEach(input => {
          const span = document.createElement('span');
          if (input.type === 'checkbox' || input.type === 'radio') {
            span.textContent = input.checked ? '☑ ' : '☐ ';
            span.style.fontWeight = 'bold';
          } else {
            span.textContent = input.value || '';
            span.style.fontWeight = '500';
          }
          input.parentNode.insertBefore(span, input);
          input.remove();
        });

        clone.querySelectorAll('.no-print, button').forEach(el => el.remove());

        // 挂载到一个精准的 210mm 隐藏容器
        const tempDiv = document.createElement('div');
        tempDiv.style.position = 'fixed';
        tempDiv.style.top = '-99999px';
        tempDiv.style.left = '0';
        tempDiv.style.width = '210mm';
        tempDiv.style.minHeight = '297mm';
        tempDiv.style.backgroundColor = '#ffffff';
        tempDiv.style.boxSizing = 'border-box';
        tempDiv.appendChild(clone);
        document.body.appendChild(tempDiv);

        try {
          const canvas = await window.html2canvas(clone, {
            scale: 2, // 2倍超清Retina
            useCORS: true,
            backgroundColor: '#ffffff',
            logging: false
          });

          const imgData = canvas.toDataURL('image/jpeg', 0.98);

          if (i > 0) {
            pdf.addPage('a4', 'p');
          }

          // 严丝合缝绘制当前页，绝对无多余空白切片
          pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight);
        } finally {
          document.body.removeChild(tempDiv);
        }
      }

      pdf.save(fileName);
      return;
    }

    // 备用兜底: 原生打印
    window.print();
  }

  return {
    exportToWord,
    exportToPdf,
    triggerNativePrint: () => window.print()
  };
})();

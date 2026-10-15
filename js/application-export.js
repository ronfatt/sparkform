/**
 * Spark One 申请表导出引擎 (Word & PDF)
 * 严格按照原版 PDF 规格排版：
 * - 办公室补贴申请表: 严格 2 页 (Page 1 + Page 2)
 * - 讲师补贴申请表: 严格 2 页 (Page 1 + Page 2)
 * - 活动申请表: 严格 1 页 (Single Page)
 * - 市场领航扶持申请表: 紧凑 A4
 */

window.ApplicationExporter = (function () {
  'use strict';

  function getTodayString() {
    const d = new Date();
    return `${d.getFullYear()}年${String(d.getMonth() + 1).padStart(2, '0')}月${String(d.getDate()).padStart(2, '0')}日`;
  }

  function getFileName(formId, ext) {
    const d = new Date().toISOString().split('T')[0];
    if (formId === 'office-subsidy') return `Spark_办公室补贴申请表_${d}.${ext}`;
    if (formId === 'lecturer-subsidy') return `Spark_讲师补贴申请表_${d}.${ext}`;
    if (formId === 'event-apply') return `Spark_活动申请表_${d}.${ext}`;
    return `Spark_市场领航扶持账号申请表_${d}.${ext}`;
  }

  // =========================================================================
  // 1. 导出 WORD (.docx)
  // =========================================================================
  async function exportToWord(formId, formData, i18nDict) {
    const dict = i18nDict || (window.APPLICATION_I18N && window.APPLICATION_I18N[window.currentAppLang || 'zh-CN']) || {};
    const t = (k, fallback) => (dict[k] !== undefined ? dict[k] : fallback);
    if (!window.docx) {
      alert('Word 导出组件加载中，请稍后再试...');
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
            new TextRun({ text: `${t('app_date_lbl', '申请日期：')} ${dateStr || getTodayString()}        `, size: 17, font: "Microsoft YaHei" }),
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
                new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: purpleHeader, margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "物料项目", bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 35, type: WidthType.PERCENTAGE }, shading: purpleHeader, margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "准备情况", bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, shading: purpleHeader, margins: { top: 40, bottom: 40, left: 80, right: 80 }, children: [new Paragraph({ children: [new TextRun({ text: "数量∕说明", bold: true, color: "FFFFFF", size: 17 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "公司标识、门牌及背景墙", size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☑已完成  ☐筹备中  ☐未准备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_1 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "公司介绍及Spark One介绍资料", size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☑已完成  ☐筹备中  ☐未准备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_2 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "经审核的课程课件及讲解资料", size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☑已完成  ☐筹备中  ☐未准备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_3 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "风险说明及客户服务指引", size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☑已完成  ☐筹备中  ☐未准备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_4 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "桌椅、投影∕显示设备", size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☑已完成  ☐筹备中  ☐未准备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_5 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "网络、收音及视频拍摄设备", size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☑已完成  ☐筹备中  ☐未准备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_6 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "签到表、活动记录及咨询登记表", size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☑已完成  ☐筹备中  ☐未准备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_7 || "", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "其他物料", size: 17 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☐已完成  ☑筹备中  ☐未准备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: formData.mat_desc_8 || "", size: 16 })] })] })] })
          ]
        }),
        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "01 / 02", size: 16, bold: true, color: "64748B" })], spacing: { before: 120 } }),
        
        // --- 强制分页: 产生第 2 页 ---
        new Paragraph({ children: [new PageBreak()] }),

        // --- PAGE 2 ---
        new Paragraph({
          children: [
            new TextRun({ text: "SPARK UNION CAPITAL INC.", bold: true, size: 20, color: "2E1065" }),
            new TextRun({ text: "                     办公室补贴申请表", bold: true, size: 24, color: "1E1B4B" })
          ],
          spacing: { after: 60 }
        }),
        createSectionBar(t('off_sec3', "三、社区群运营情况")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow("社区群名称及平台", formData.off_comm_name),
            createTwoColRow("社区群链接∕群编号", formData.off_comm_link),
            createTwoColRow("群主及管理员", formData.off_comm_admins),
            createTwoColRow("当前群人数", formData.off_comm_members),
            createTwoColRow("近30天活跃人数及统计口径", formData.off_comm_active),
            createTwoColRow("服务地区及主要语言", formData.off_comm_lang),
            createTwoColRow("每日内容发布安排", formData.off_comm_daily),
            createTwoColRow("每周课程及活动安排", formData.off_comm_weekly),
            createTwoColRow("客户问题反馈与处理方式", formData.off_comm_feedback)
          ]
        }),
        new Paragraph({ spacing: { before: 80, after: 40 } }),
        createSectionBar(t('off_sec4', "四、每日讲课视频提交要求")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Paragraph({ children: [new TextRun({ text: "1. 提交频次：每个批准的运营日提交不少于 1 条当日实际授课视频；休息日及停课情况提前报备。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "2. 视频时长：每条不少于 15 分钟；☑完整课程  ☐课堂片段＋完整录像留档。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "3. 视频内容：注明日期、办公室名称、讲师姓名、课程主题及参与人数，清晰展示实际授课情况。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "4. 课程范围：公司介绍、产品规则、平台操作、风险教育及客户服务；使用经审核的资料。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "5. 提交方式：每日 21:00 前提交至官方审核群，同步附上签到人数及课程简报。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "6. 真实性要求：不得使用旧视频重复提交、拼接伪造授课记录；拍摄参与者前取得同意。", size: 16 })], spacing: { after: 80 } }),
        createSectionBar(t('off_sec5', "五、考核期限与目标")),
        new Paragraph({
          children: [
            new TextRun({ text: `考核期限：${formData.off_kpi_period || '______年____月____日 至 ______年____月____日'}        `, size: 16 }),
            new TextRun({ text: "考核频率：☐每周  ☑每月  ☐其他", size: 16 })
          ],
          spacing: { before: 40, after: 30 }
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 34, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "考核指标", bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "批准目标", bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "实际完成", bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "审核结果", bold: true, color: "FFFFFF", size: 17 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "办公室实际开放天数", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 24 天", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "课程∕活动场次", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 8 场", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "有效视频提交天数", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 20 天", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "课程参与人数∕人次", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 150 人次", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "社区活跃人数", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 100 人", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "客户咨询及问题处理完成率", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 98%", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "业绩目标：", bold: true, size: 16 })] })] }), new TableCell({ columnSpan: 3, children: [new Paragraph({ children: [new TextRun({ text: formData.off_kpi_goal || "直推量化托管额度突破 200,000 美元", size: 16 })] })] })] })
          ]
        }),
        new Paragraph({
          children: [new TextRun({ text: "附件清单： ☑租赁合同  ☑租金支付凭证  ☑办公室现场照片  ☑物料照片  ☑社区群信息截图  ☑课程计划  ☐其他", size: 15, color: "334155" })],
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
            createTwoColRow("讲师姓名", formData.lec_name),
            createTwoColRow("会员编号∕职务", formData.lec_account),
            createTwoColRow("所属国家∕城市", formData.lec_city),
            createTwoColRow("联系电话∕社交账号", formData.lec_contact),
            createTwoColRow("授课语言", formData.lec_lang),
            createTwoColRow("所属团队∕办公室", formData.lec_team),
            createTwoColRow("推荐负责人及联系方式", formData.lec_ref),
            createTwoColRow("授课经验及擅长主题", formData.lec_exp),
            createTwoColRow("申请补贴金额及币种", formData.lec_amt),
            createTwoColRow("申请补贴方式", "☑按半月   ☐按场   ☐其他")
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
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "社区名称", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 18, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "平台∕群编号", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 16, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "服务地区", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 12, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "群人数", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 16, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "负责人", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 16, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "计划频次", bold: true, color: "FFFFFF", size: 16 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "亚太先锋财富群", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "TG官方群", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "新加坡/大马", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "500人", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "张总", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "每周2次", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "华南量化私享群", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "微信群08", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "深圳/香港", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "380人", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "刘经理", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "每周3次", size: 16 })] })] })] })
          ]
        }),
        new Paragraph({ children: [new TextRun({ text: `主要服务对象及需求：${formData.lec_audience || "高净值投资者与新加入量化会员"}`, size: 16 })], spacing: { before: 40, after: 60 } }),
        createSectionBar(t('lec_sec3', "三、授课计划")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "课程方向", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "拟讲主题", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 15, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "计划场次", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 15, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "单场时长", bold: true, color: "FFFFFF", size: 16 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "公司介绍与发展规划", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "企业愿景与MSB合规架构", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "4场", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "45分钟", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Spark One 产品及规则说明", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "黄金与加密量化稳健收益拆解", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "6场", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "60分钟", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "平台操作与客户服务", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "DApp工作台充提与质押实操", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "4场", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "30分钟", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "投资风险及账户安全教育", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "私钥保管与防钓鱼诈骗教育", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "2场", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "40分钟", size: 16 })] })] })] })
          ]
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "课件版本/审核人：V5.2 / 教育部审核组        ", size: 16 }),
            new TextRun({ text: "授课形式：☑线下  ☑线上直播  ☐录播  ☐混合", size: 16 })
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
            new TextRun({ text: "                     讲师补贴申请表", bold: true, size: 24, color: "1E1B4B" })
          ],
          spacing: { after: 60 }
        }),
        createSectionBar(t('lec_sec4', "四、每日视频规则")),
        new Paragraph({ spacing: { before: 30, after: 30 } }),
        new Paragraph({ children: [new TextRun({ text: "1. 提交频次：每个批准的授课日提交不少于 1 条本人当日授课视频。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "2. 视频时长：每条不少于 15 分钟；☑完整课程  ☐课程片段+完整录像留档。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "3. 视频标注：讲师姓名、授课日期、服务社区、课程主题及实际参与人数。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "4. 质量要求：画面与声音清晰，内容完整连贯，能够核验实际授课情况。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "5. 提交时间：每日 22:00 前提交至官方讲师专群，附课程要点、参与人数及问题反馈。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "6. 内容要求：使用经审核的课件，准确说明产品规则及风险；不得承诺保本收益。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "7. 记录要求：不得以旧视频冒充当日课程；同一场课程服务多个社区时须注明。", size: 16 })] }),
        new Paragraph({ children: [new TextRun({ text: "8. 隐私要求：公开使用含学员面部、姓名或联系方式的视频前，取得相应同意。", size: 16 })], spacing: { after: 60 } }),
        createSectionBar(t('lec_sec5', "五、考核期限与业绩目标")),
        new Paragraph({
          children: [
            new TextRun({ text: `考核期限：${formData.lec_kpi_period || '______年____月____日 至 ______年____月____日'}        `, size: 16 }),
            new TextRun({ text: "考核频率：☐每周  ☑每月  ☐其他", size: 16 })
          ],
          spacing: { before: 30, after: 30 }
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 34, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "考核指标", bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "批准目标", bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "实际完成", bold: true, color: "FFFFFF", size: 17 })] })] }),
                new TableCell({ width: { size: 22, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "审核结果", bold: true, color: "FFFFFF", size: 17 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "实际授课天数", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 20 天", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "有效授课场次", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 16 场", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "有效视频提交率", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "100%", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "服务社区数量", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 2 个", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "课程参与人数∕人次", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 300 人次", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "学员满意度", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "≥ 95%", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "课件准确性及讲解质量", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "优秀", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合格", size: 16 })] })] })] })
          ]
        }),
        new Paragraph({ spacing: { before: 60, after: 30 } }),
        createSectionBar(t('lec_sec6', "六、附件及申请人确认")),
        new Paragraph({
          children: [new TextRun({ text: "附件： ☑个人简介  ☑试讲视频  ☑课程计划  ☑课件样本  ☑服务社区证明  ☐其他", size: 15 })],
          spacing: { before: 30, after: 30 }
        }),
        new Paragraph({
          children: [new TextRun({ text: "申请人确认：本人确认申请资料真实，按批准计划授课并提交记录，遵守课件审核及宣传要求。无法按计划授课时，提前向负责人报备。", size: 15, color: "334155" })],
          spacing: { after: 50 }
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            new TableRow({
              children: [
                new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: "申请人签名", bold: true, size: 16 })] })] }),
                new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: formData.lec_sign || "张峰", size: 16 })] })] }),
                new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: "推荐负责人", bold: true, size: 16 })] })] }),
                new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: formData.lec_ref_sign || "陈志远", size: 16 })] })] })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({ shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: "审批结果", bold: true, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "☑批准  ☐需修改  ☐不批准", size: 15 })] })] }),
                new TableCell({ shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: "审核负责人", bold: true, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "官方审核组", size: 16 })] })] })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({ shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: "审核日期", bold: true, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: getTodayString(), size: 16 })] })] }),
                new TableCell({ shading: grayBg, children: [new Paragraph({ children: [new TextRun({ text: "备注", bold: true, size: 16 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "予以全额批准", size: 16 })] })] })
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
            new TextRun({ text: "                     活动申请表", bold: true, size: 26, color: "1E1B4B" })
          ],
          spacing: { after: 60 }
        }),
        createSectionBar(t('act_sec1', "1. 活动基本信息")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow("活动名称", formData.act_name, 22),
            createTwoColRow("活动类型", "☑会员交流会  ☑财富分享会  ☐新人说明会  ☐推广活动  ☐社区沙龙  ☐公益活动", 22),
            createTwoColRow("活动日期", formData.act_date || getTodayString(), 22),
            createTwoColRow("活动地点", formData.act_venue, 22),
            createTwoColRow("预计参与人数", "☐<10人  ☐10-30人  ☑30-50人  ☐50-100人  ☐100人以上", 22)
          ]
        }),
        new Paragraph({ spacing: { before: 40, after: 20 } }),
        createSectionBar(t('act_sec2', "2. 申请人信息")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow("申请人姓名", formData.act_app_name, 22),
            createTwoColRow("联系电话", formData.act_app_phone, 22),
            createTwoColRow("电子邮箱", formData.act_app_email, 22),
            createTwoColRow("所属市场团队", formData.act_app_team, 22)
          ]
        }),
        new Paragraph({ spacing: { before: 40, after: 20 } }),
        createSectionBar(t('act_sec3', "3. 活动内容说明")),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow("活动主题", formData.act_theme, 22),
            createTwoColRow("活动目的", "☑品牌宣传  ☑新会员招募  ☑市场拓展  ☑产品介绍  ☐投资知识分享  ☐团队建设", 22),
            createTwoColRow("活动流程简介", "1. 嘉宾签到与宣传片暖场 (30分)\n2. Spark One 量化与合规架构介绍 (45分)\n3. 盘口实操与答疑 (30分)\n4. VIP商务茶歇交流 (45分)", 22)
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
                new TableCell({ width: { size: 60, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "项目", bold: true, color: "FFFFFF", size: 16 })] })] }),
                new TableCell({ width: { size: 40, type: WidthType.PERCENTAGE }, shading: purpleHeader, children: [new Paragraph({ children: [new TextRun({ text: "预算金额 (USD)", bold: true, color: "FFFFFF", size: 16 })] })] })
              ]
            }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "场地租金及音响设备", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `$${formData.act_b1 || '1200'}`, size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "餐饮茶歇及果盘饮品", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `$${formData.act_b2 || '600'}`, size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "宣传物料、易拉宝及伴手礼", size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `$${formData.act_b3 || '400'}`, size: 16 })] })] })] }),
            new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "合计", bold: true, size: 16 })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `$${formData.act_btotal || '2200'} USD`, bold: true, size: 16, color: "DC2626" })] })] })] })
          ]
        }),
        new Paragraph({
          children: [new TextRun({ text: "5. 物料支持： ☑企业宣传册  ☑品牌易拉宝  ☐活动海报  ☑宣传视频  ☑产品介绍PPT  ☐礼品支持", size: 15 })],
          spacing: { before: 30, after: 30 }
        }),
        new Paragraph({
          children: [new TextRun({ text: `6. 活动负责人： 第一负责人：${formData.act_resp1 || '李立明 13800000001'}      第二负责人：${formData.act_resp2 || '张峰 13900000002'}`, size: 15 })],
          spacing: { after: 30 }
        }),
        new Paragraph({
          children: [new TextRun({ text: "7. 风险管理承诺： 严格遵守当地法律法规；不进行虚假宣传；不承诺保本固定收益；维护星火联合资本品牌形象。", size: 14, color: "475569" })],
          spacing: { after: 30 }
        }),
        new Paragraph({
          children: [
            new TextRun({ text: `申请人签字： ${formData.act_sign || '李立明'}                     `, bold: true, size: 16 }),
            new TextRun({ text: `日期： ${formData.act_sign_date || getTodayString()}`, bold: true, size: 16 })
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
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `实际人数: ${formData.act_p_num || '48'} 人`, size: 15 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `新注册: ${formData.act_reg_num || '22'} 人`, size: 15 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `新投资: ${formData.act_inv_num || '15'} 人`, size: 15 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `照片: ${formData.act_pic_num || '85'} 张`, size: 15 })] })] })
              ]
            }),
            new TableRow({
              children: [
                new TableCell({ columnSpan: 4, children: [new Paragraph({ children: [new TextRun({ text: `活动总结: ${formData.act_summary || '活动圆满成功，现场意向签约超38万美元，全面达成预期指标。'}`, size: 15 })] })] })
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
            new TextRun({ text: "           市场领航扶持账号额度申请表", bold: true, size: 26, color: "1E1B4B" })
          ],
          spacing: { after: 60 }
        }),
        createSectionBar("一、申请人基本资料"),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow("姓名", formData.sec1_name, 25),
            createTwoColRow("会员账号", formData.sec1_account, 25),
            createTwoColRow("联系电话/邮箱", formData.sec1_contact, 25),
            createTwoColRow("常驻国家/城市", formData.sec1_residence, 25),
            createTwoColRow("每周投入时间", `${formData.sec1_hours || '35'} 小时/周`, 25),
            createTwoColRow("业绩情况", formData.sec1_perf_detail, 25)
          ]
        }),
        new Paragraph({ spacing: { before: 30, after: 20 } }),
        createSectionBar("二、申请扶持额度与规划"),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow("申请额度", "$10,000 美元", 25),
            createTwoColRow("目标开拓城市", formData.sec2_regions, 25),
            createTwoColRow("理由及规划", formData.sec2_reason, 25)
          ]
        }),
        new Paragraph({ spacing: { before: 30, after: 20 } }),
        createSectionBar("三、30天预期目标与声明"),
        new Paragraph({ spacing: { before: 20, after: 20 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE }, borders: tableBorders,
          rows: [
            createTwoColRow("触达客户人数", formData.sec4_target_clients, 25),
            createTwoColRow("举办说明会场次", formData.sec4_target_events, 25),
            createTwoColRow("合规承诺与签署", `本人确认信息真实，未达10倍业绩自觉接受收回。\n申请人签署：${formData.sec1_name || '陈志远'}        日期：${getTodayString()}`, 25)
          ]
        })
      );
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 600, right: 600, bottom: 600, left: 600 } // 精准窄边距，确保完美分页
          }
        },
        children: docChildren
      }]
    });

    const blob = await Packer.toBlob(doc);
    const fileName = getFileName(formId, 'docx');
    window.saveAs(blob, fileName);
    return fileName;
  }

  // =========================================================================
  // 2. 导出 PDF (.pdf) - 严格按 .a4-paper 逐页 1:1 独立绘制，彻底根除任何多余空白页！
  // =========================================================================
  async function exportToPdf(formId, targetElement) {
    const fileName = getFileName(formId, 'pdf');
    
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

const pptxgen = require('pptxgenjs');

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { name, contact, area, target, exp_type, exp_units, exp_days, exp_leads, lang } = req.body;

  if (!name || !contact || !area) {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  const isZh = lang === 'zh';
  const isOp = target === 'operator';

  const BG='0D0D0D', GOLD='D4A843', GOLD2='C8A96E', WHITE='E8E0D0';
  const GREY='706050', DK='504840', CARD='111111', CARD_B='2A2A2A';
  const HLBG='1A1608', FONT='Arial';

  let expLine, expBadge;
  if (exp_type==='A') {
    expLine  = isZh ? `已成功协助 ${exp_units} 间房间出租，最快 ${exp_days} 天完成` : `Successfully helped ${exp_units} rooms get rented, fastest in ${exp_days} days`;
    expBadge = isZh ? `${exp_units} 间成功出租` : `${exp_units} Rooms Rented`;
  } else if (exp_type==='B') {
    expLine  = isZh ? `手上现有 ${exp_leads} 位潜在租客正积极寻找 ${area} 房间，可即时推介` : `Currently holding ${exp_leads} active leads seeking rooms in ${area}, ready to introduce immediately`;
    expBadge = isZh ? `${exp_leads} 位潜在租客在手` : `${exp_leads} Active Leads Ready`;
  } else {
    expLine  = isZh ? `专注 ${area} 区域出租推广，认真跟进每一个 lead，确保屋主获得高质量租客` : `Focused on ${area} rental marketing, committed to following up every lead`;
    expBadge = isZh ? '专注执行 · 认真负责' : 'Committed & Focused';
  }

  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';

  function sl(cb) { const s=pres.addSlide(); s.background={color:BG}; cb(s); }
  function lbl(s,t,y){ s.addText(t,{x:0.6,y,w:9,h:0.22,fontSize:9,color:DK,charSpacing:4,fontFace:FONT}); }
  function ttl(s,t,y){ s.addText(t,{x:0.6,y,w:9,h:0.52,fontSize:28,bold:true,color:GOLD,fontFace:FONT}); }

  // SLIDE 1: Cover
  sl(s => {
    s.addShape('RECTANGLE',{x:0,y:0,w:10,h:0.06,fill:{color:'1A1608'},line:{color:'1A1608'}});
    s.addText(isZh?'Lead Provider · 出租推广服务':'Lead Provider · Rental Marketing Services',
      {x:0.5,y:0.55,w:9,h:0.35,align:'center',fontSize:13,color:DK,charSpacing:3,fontFace:FONT});
    s.addText('PROPOSAL',{x:0.52,y:1.07,w:9,h:1.8,align:'center',fontSize:110,bold:true,color:GOLD,fontFace:FONT});
    s.addShape('RECTANGLE',{x:3.2,y:2.95,w:3.6,h:0.015,fill:{color:'2A2A2A'},line:{color:'2A2A2A'}});
    s.addText(name,{x:0.5,y:3.08,w:9,h:0.55,align:'center',fontSize:26,bold:true,color:WHITE,fontFace:FONT});
    s.addText(isZh?(isOp?'针对 Operator':'针对屋主'):(isOp?'For Operators':'For Landlords'),
      {x:0.5,y:3.68,w:9,h:0.3,align:'center',fontSize:14,color:GOLD2,charSpacing:3,fontFace:FONT});
    s.addShape('RECTANGLE',{x:3.5,y:4.1,w:3.0,h:0.015,fill:{color:'1E1E1E'},line:{color:'1E1E1E'}});
    s.addText(`WhatsApp: ${contact}`,{x:0.5,y:4.22,w:9,h:0.28,align:'center',fontSize:12,color:'404838',fontFace:FONT});
    s.addText('Presentation · 2026',{x:0.5,y:4.52,w:9,h:0.22,align:'center',fontSize:10,color:'252520',charSpacing:3,fontFace:FONT});
  });

  // SLIDE 2: About + Services
  sl(s => {
    lbl(s,isZh?'01 / 关于我':'01 / ABOUT ME',0.28);
    ttl(s,isZh?'我是谁':'Who I Am',0.52);
    s.addText(isZh?`我是 ${name}，专注于 ${area} 一带的租房市场推广服务，为有出租需求的${isOp?'Operator':'屋主'}提供高质量的租客线索`:`I am ${name}, specialising in rental marketing in ${area}, providing high-quality tenant leads for ${isOp?'operators':'landlords'} who need to rent out their properties`,
      {x:0.6,y:1.1,w:8.8,h:0.4,fontSize:12.5,color:'B0A898',fontFace:FONT,wrap:true});
    s.addText(expLine,{x:0.6,y:1.55,w:8.8,h:0.35,fontSize:12.5,color:'B0A898',fontFace:FONT,wrap:true});
    s.addShape('RECTANGLE',{x:0.6,y:2.0,w:3.2,h:0.38,fill:{color:HLBG},line:{color:GOLD,width:0.8}});
    s.addText(expBadge,{x:0.6,y:2.0,w:3.2,h:0.38,align:'center',fontSize:13,bold:true,color:GOLD,fontFace:FONT});
    s.addShape('RECTANGLE',{x:0.6,y:2.55,w:8.8,h:0.012,fill:{color:'1E1E1E'},line:{color:'1E1E1E'}});
    lbl(s,isZh?'02 / 我的服务':'02 / MY SERVICES',2.68);
    const svcs=[
      {t:isZh?'多平台发布 Posting':'Multi-Platform Posting',d:isZh?'在 Facebook 群组、iBilik 等主要平台发布房间，覆盖最多潜在租客':'Post on Facebook Groups, iBilik and more to reach maximum potential tenants'},
      {t:isZh?'筛选 Quality Lead':'Quality Lead Screening',d:isZh?'主动筛选有预算、有搬入时间、准备好看房的租客，不是随便转介联系方式':'Screen tenants with budget, move-in date and ready to view — not random enquiries'},
      {t:isZh?'安排 Viewing 时间':'Viewing Arrangement',d:isZh?'协助约好看房时间，减少你的时间成本，只需接待真正认真的租客':'Schedule viewings to save your time — you only meet serious tenants'},
      {t:isZh?'全程跟进':'End-to-End Follow-Up',d:isZh?'从发 posting 到租客入住，全程跟进，确保每个 lead 被认真对待':'From posting to move-in, follow up every step so no lead is left behind'},
    ];
    const cols=[0.6,5.2];
    svcs.forEach((sv,i)=>{
      const x=cols[i%2],y=2.95+Math.floor(i/2)*1.28;
      s.addShape('RECTANGLE',{x,y,w:4.4,h:1.15,fill:{color:CARD},line:{color:CARD_B,width:0.5}});
      s.addText(sv.t,{x:x+0.15,y:y+0.1,w:4.1,h:0.28,fontSize:12.5,bold:true,color:GOLD2,fontFace:FONT});
      s.addText(sv.d,{x:x+0.15,y:y+0.42,w:4.1,h:0.65,fontSize:11,color:GREY,fontFace:FONT,wrap:true});
    });
  });

  // SLIDE 3: Why Choose
  sl(s => {
    lbl(s,isZh?'03 / 为什么选择我':'03 / WHY CHOOSE ME',0.28);
    ttl(s,isZh?'合作对你的价值':'Value for You',0.52);
    const benefits=[
      {t:isZh?'只需提供房间资料和照片，其余全由我负责':'Just provide room details and photos — I handle everything else',d:isZh?'推广、筛选、跟进，全部我来':'Marketing, screening and follow-up, all handled by me'},
      {t:isZh?'零前期费用，完全零风险':'Zero upfront fees, completely risk-free',d:isZh?'不成功不收钱，你没有任何损失':'No success, no charge — nothing to lose'},
      {t:isZh?'只推介 Quality Lead，不浪费你的时间':'Only Quality Leads — no wasted viewings',d:isZh?'有预算、有搬入时间、愿意看房才会转介':'Budget, move-in date and viewing availability confirmed before referral'},
      {t:isZh?'成功才付佣金，利益一致':'Commission only upon success — aligned interests',d:isZh?'我的收入来自你的成功，所以我会全力以赴':'My income depends on your success — I give my best every time'},
    ];
    benefits.forEach((b,i)=>{
      const y=1.2+i*1.04;
      s.addShape('RECTANGLE',{x:0.6,y,w:8.8,h:0.88,fill:{color:CARD},line:{color:'1E1E1E',width:0.5}});
      s.addText('✔',{x:0.75,y:y+0.12,w:0.4,h:0.32,fontSize:15,color:GOLD,fontFace:FONT});
      s.addText(b.t,{x:1.22,y:y+0.1,w:8.0,h:0.3,fontSize:13.5,bold:true,color:WHITE,fontFace:FONT});
      s.addText(b.d,{x:1.22,y:y+0.45,w:8.0,h:0.28,fontSize:11.5,color:GREY,fontFace:FONT});
    });
  });

  // SLIDE 4: Commission
  sl(s => {
    lbl(s,isZh?'04 / 佣金条件':'04 / COMMISSION TERMS',0.28);
    ttl(s,isZh?'透明的收费结构':'Transparent Fee Structure',0.52);
    if (isOp) {
      [{amount:'RM20',label:isZh?'每个确认看房的\nQuality Lead':'Per confirmed\nviewing lead',hi:false},
       {amount:'20%',label:isZh?'租客成功入住后\n首月租金佣金':'Of first month\'s rent\nupon successful move-in',hi:true}
      ].forEach((c,i)=>{
        const x=i===0?0.6:5.3,bg=c.hi?HLBG:CARD,bc=c.hi?GOLD:CARD_B;
        s.addShape('RECTANGLE',{x,y:1.25,w:4.3,h:2.7,fill:{color:bg},line:{color:bc,width:c.hi?1:0.5}});
        s.addText(c.amount,{x,y:1.55,w:4.3,h:1.2,align:'center',fontSize:60,bold:true,color:GOLD,fontFace:FONT});
        s.addText(c.label,{x:x+0.2,y:2.85,w:3.9,h:0.8,align:'center',fontSize:13.5,color:GREY,fontFace:FONT,wrap:true});
      });
      s.addShape('RECTANGLE',{x:0.6,y:4.15,w:8.8,h:0.52,fill:{color:CARD},line:{color:'1E1E1E',width:0.5}});
      s.addText(isZh?'不成功不收费——租客未入住前，你无需支付任何费用':'No success, no fee — you pay nothing until the tenant successfully moves in',
        {x:0.7,y:4.22,w:8.6,h:0.35,align:'center',fontSize:12.5,color:GREY,fontFace:FONT});
    } else {
      [{amount:'25%',label:isZh?'租约 3 个月\n佣金':'3-Month Tenancy\nCommission',hi:false},
       {amount:'50%',label:isZh?'租约 6 个月\n佣金':'6-Month Tenancy\nCommission',hi:false},
       {amount:'100%',label:isZh?'租约 12 个月\n佣金\n（市场标准）':'12-Month Tenancy\nCommission\n(Market Rate)',hi:true}
      ].forEach((c,i)=>{
        const x=0.5+i*3.1,bg=c.hi?HLBG:CARD,bc=c.hi?GOLD:CARD_B;
        s.addShape('RECTANGLE',{x,y:1.25,w:2.9,h:2.85,fill:{color:bg},line:{color:bc,width:c.hi?1:0.5}});
        s.addText(c.amount,{x,y:1.5,w:2.9,h:1.0,align:'center',fontSize:50,bold:true,color:GOLD,fontFace:FONT});
        s.addText(c.label,{x:x+0.1,y:2.62,w:2.7,h:1.2,align:'center',fontSize:13,color:GREY,fontFace:FONT,wrap:true});
      });
      s.addShape('RECTANGLE',{x:0.5,y:4.28,w:9.2,h:0.52,fill:{color:CARD},line:{color:'1E1E1E',width:0.5}});
      s.addText(isZh?'佣金标准与市场中介一致，透明合理，成功才结算':'Commission aligned with market agent rates — transparent, fair, success-based',
        {x:0.6,y:4.35,w:9.0,h:0.35,align:'center',fontSize:12.5,color:GREY,fontFace:FONT});
    }
  });

  // SLIDE 5: Process
  sl(s => {
    lbl(s,isZh?'05 / 合作流程':'05 / HOW IT WORKS',0.28);
    ttl(s,isZh?'怎么开始合作':'Our 4-Step Process',0.52);
    const steps=[
      {t:isZh?'你提供房间资料 + 照片':'You provide room details + photos',d:isZh?'房间详情、月租、包含设施，以及清晰的房间照片':'Room info, monthly rent, facilities and clear room photos'},
      {t:isZh?'我负责发 Posting + 筛选租客':'I handle posting + tenant screening',d:isZh?'在各大平台发布，主动筛选符合条件的 quality lead':'Publish on all platforms, actively screen for quality leads'},
      {t:isZh?'有 Quality Lead 才通知你安排 Viewing':'Only notify you when a Quality Lead is ready',d:isZh?'确认租客有预算、有搬入时间、愿意看房，才联系你':'Tenant confirmed with budget, move-in date and willing to view before contacting you'},
      {t:isZh?'租客成功入住，佣金才结算':'Tenant moves in — commission is settled',d:isZh?'入住后才结算，未入住前你无需支付任何费用':'Commission settled upon move-in — no charges before that'},
    ];
    steps.forEach((st,i)=>{
      const y=1.2+i*1.02;
      s.addShape('RECTANGLE',{x:0.6,y,w:8.8,h:0.86,fill:{color:CARD},line:{color:'1E1E1E',width:0.5}});
      s.addShape('RECTANGLE',{x:0.72,y:y+0.22,w:0.34,h:0.34,fill:{color:HLBG},line:{color:GOLD,width:0.8}});
      s.addText(`${i+1}`,{x:0.72,y:y+0.22,w:0.34,h:0.34,align:'center',fontSize:12,bold:true,color:GOLD,fontFace:FONT});
      s.addText(st.t,{x:1.2,y:y+0.1,w:8.0,h:0.3,fontSize:13.5,bold:true,color:WHITE,fontFace:FONT});
      s.addText(st.d,{x:1.2,y:y+0.48,w:8.0,h:0.28,fontSize:11.5,color:GREY,fontFace:FONT});
    });
  });

  // SLIDE 6: Contact
  sl(s => {
    lbl(s,isZh?'06 / 联系我':'06 / CONTACT',0.28);
    ttl(s,isZh?'开始合作':"Let's Work Together",0.52);
    s.addShape('RECTANGLE',{x:0.6,y:1.22,w:8.8,h:2.5,fill:{color:CARD},line:{color:CARD_B,width:0.5}});
    [[isZh?'姓名':'NAME',name,WHITE],['WHATSAPP',contact,GOLD2],[isZh?'服务地区':'SERVICE AREA',area,WHITE]]
    .forEach(([l,v,c],i)=>{
      const y=1.45+i*0.72;
      s.addText(l,{x:0.85,y,w:1.4,h:0.28,fontSize:9,color:DK,charSpacing:2,fontFace:FONT});
      s.addText(v,{x:2.5,y,w:6.7,h:0.3,fontSize:15,bold:true,color:c,fontFace:FONT});
    });
    s.addText(isZh?'如有合作意向，欢迎随时 WhatsApp 联系我':'Interested? Feel free to WhatsApp me anytime',
      {x:0.6,y:3.9,w:8.8,h:0.35,align:'center',fontSize:13,color:GREY,fontFace:FONT});
    s.addText(isZh?'期待与你合作，一起把房间出租出去 🏠':'Looking forward to partnering with you 🏠',
      {x:0.6,y:4.28,w:8.8,h:0.35,align:'center',fontSize:14,bold:true,color:GOLD2,fontFace:FONT});
  });

  try {
    const buffer = await pres.write({ outputType: 'nodebuffer' });
    const filename = `Proposal_${name.replace(/[^a-zA-Z0-9]/g,'_')}_${target}_${lang}.pptx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: 'Failed to generate file' });
  }
};

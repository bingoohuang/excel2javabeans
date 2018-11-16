package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.beans2excel.CepingResult.ItemComment;
import com.github.bingoohuang.excel2beans.BeansToExcelOnTemplate;
import com.github.bingoohuang.excel2beans.PoiUtil;
import com.github.bingoohuang.utils.lang.Collects;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import static com.github.bingoohuang.beans2excel.CepingResult.Item;

public class BeansToExcelOnTemplateTest {
    @Test
    public void test2() {
        val builder = CepingResult.builder();
        builder
                .sheetName("王雷测评结论表")
                .interviewCode("20181101.001")
                .name("王雷")
                .gender("男")
                .age("38")
                .position("会计")
                .level("18-24级")
                .annualSalary("18")
                .matchScore(2.36);

        builder
                .itemComment(ItemComment.builder().item("优势").comment("使命认同:").build())
                .itemComment(ItemComment.builder().item("优势").comment("责任心:愿意接纳自己的一些小错误,对工作抱以轻松的态度,容易承担不必要的、负面的风险,担心因犯错而被批评，不愿意独立行动或决定,偏好制定自己的规则。").build())
                .itemComment(ItemComment.builder().item("优势").comment("学习能力:对事物工作原理不感兴趣,消极地对待教育体验,有点健忘,容易分心，注意力不集中,知识过时，不更新。").build())
                .itemComment(ItemComment.builder().item("优势").comment("坚韧不拔:对压力非常敏感，容易受到压力的困扰,容易放弃。").build())
                .itemComment(ItemComment.builder().item("优势").comment("目标结果导向:对所获的成就不满意,缺乏清晰明确的信念或兴趣,回避竞争，缺乏韧性,不自信，有些自我怀疑,过于关注细节而忽略大局。").build())
                .itemComment(ItemComment.builder().item("优势").comment("执行力:计划性差，较为随机,对于规则与程序持有较为随意的态度,不时会情绪波动,缺乏职业发展方向,过度依赖于他人的建议，而不愿独立决断或行动。").build())
                .itemComment(ItemComment.builder().item("优势").comment("团队合作:质疑他人的动机,偏好小型群体,并不总是能和他人友好相处,容易被他人的过错激怒,对于人和事普遍存疑。").build())
                .itemComment(ItemComment.builder().item("优势").comment("勤奋:对工作成果较为随意，容易满足现状,对工作抱以轻松的态度,过于关注细节而忽略大局,心态不够积极,对上级高绩效的要求表现出不满或抵触。").build())
                .itemComment(ItemComment.builder().item("优势").comment("正能量:遇到困难时心态不够积极,非常不自信。").build())
                .itemComment(ItemComment.builder().item("优势").comment("诚信:称许性正常。").build())

                .itemComment(ItemComment.builder().item("心理健康").comment("无高风险").build())
                .itemComment(ItemComment.builder().item("待提升").comment("领导潜力:不强求成败,对问题的研究不够深入,考虑问题思路不够开阔,较少带领大家一起做事。").build())

        ;

        builder
                .item(Item.builder().category("核心素质").quality("使命认同").dimension("文化价值观匹配").score("6.8").scoreTmpl("FAIL").remark("工作场景偏好").build())
                .item(Item.builder().category("基本素质").quality("心理健康").dimension("焦虑不安,抑郁消沉,偏执多疑,冷漠孤僻,特立独行,冲动暴躁,喜怒无常,社交回避,僵化固执,依赖顺从,夸张做作,狂妄自恋").score("无高风险").remark("").build())

        ;

        export(builder);
    }

    @Test
    public void test1() {
        val builder = CepingResult.builder();
        builder
                .sheetName("东方不败测评结论表")
                .interviewCode("20181101.001")
                .name("东方不败")
                .gender("不男不女")
                .age("36")
                .position("高级HR主管")
                .level("16级")
                .annualSalary("18")
                .matchScore(3.8)

                .item(Item.builder().category("基本素质").quality("心理健康").dimension("焦虑不安,抑郁消沉,偏执多疑,冷漠孤僻,特立独行,冲动暴躁,喜怒无常,社交回避,僵化固执,依赖顺从,夸张做作,狂妄自恋").score("无高风险").remark("").build())
                .item(Item.builder().category("基本素质").quality("诚信").dimension("称许性").score("正常").remark("").build())
                .item(Item.builder().category("核心素质").quality("使命认同").dimension("文化价值观匹配").score("6.8").remark("工作场景偏好(一票否决)").build())
                .item(Item.builder().category("核心素质").quality("责任心").dimension("尽责").score("1^2").remark("Hogan").build())
                .item(Item.builder().category("核心素质").quality("责任心").dimension("精通掌握").score("1^2").remark("Hogan").build())
                .item(Item.builder().category("核心素质").quality("责任心").dimension("避免麻烦").score("1^2").remark("Hogan").build())
                .item(Item.builder().category("核心素质").quality("责任心").dimension("胆怯").score("1^2").remark("Hogan").build())
                .item(Item.builder().category("核心素质").quality("责任心").dimension("循规蹈矩").score("1^2").remark("Hogan").build())
                .item(Item.builder().category("核心素质").quality("学习能力").dimension("科学能力").score("2^2").remark("Hogan").build());


        export(builder);


        builder.matchComment("观自在菩萨，行深般若波罗蜜多时，照见五蕴皆空，度一切苦厄。舍利子，色不异空，空不异色，色即是空，空即是色，受想行识，亦复如是。" +
                "舍利子，是诸法空相，不生不灭，不垢不净，不增不减。" +
                "是故空中无色，无受想行识，无眼耳鼻舌身意，无色声香味触法，无眼界，乃至无意识界，无无明，亦无无明尽，乃至无老死，亦无老死尽。" +
                "无苦集灭道，无智亦无得。以无所得故。菩提萨埵，依般若波罗蜜多故，心无挂碍。无挂碍故，无有恐怖，远离颠倒梦想，究竟涅槃。" +
                "三世诸佛，依般若波罗蜜多故，得阿耨多罗三藐三菩提。故知般若波罗蜜多，是大神咒，是大明咒，是无上咒，是无等等咒，能除一切苦，真实不虚。" +
                "故说般若波罗蜜多咒，即说咒曰：揭谛揭谛，波罗揭谛，波罗僧揭谛，菩提萨婆诃。");

        export(builder);

        builder
                .itemComment(ItemComment.builder().item("心理健康").comment("心理承受能力极强，非常乐观").build())
                .itemComment(ItemComment.builder().item("优势").comment("巴拉巴拉优势1").build())
                .itemComment(ItemComment.builder().item("优势").comment("巴拉巴拉优势2").build());

        export(builder);
    }

    @SneakyThrows
    private void export(CepingResult.CepingResultBuilder builder) {
        @Cleanup val wb = PoiUtil.getClassPathWorkbook("ceping.xlsx");
        val bean = builder.build();

        val templateName = Collects.isEmpty(bean.getItemComments())
                ? StringUtils.isEmpty(bean.getMatchComment()) ? "无总评语-模板" : "无评语-模板" : "有评语-模板";

        System.out.println(templateName);

        val beansToExcel = new BeansToExcelOnTemplate(wb.getSheet(templateName));

        @Cleanup val newWb = beansToExcel.create(bean);
        PoiUtil.protectWorkbook(newWb, "123456");

        PoiUtil.writeExcel(newWb, templateName + "-" + bean.getName() + ".xlsx");
    }
}

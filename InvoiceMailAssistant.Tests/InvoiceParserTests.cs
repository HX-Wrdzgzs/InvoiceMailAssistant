using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class InvoiceParserTests
{
    [Fact]
    public void ParsesProvidedSample()
    {
        const string body = """
            中外运向您提交了开票申请，申请信息如下：
            公司名称：测试物流有限公司
            信用代码：TEST-CREDIT-20260808
            申请金额 ：300.0 元
            申请时间：2026-08-07 10:28
            开票方式：电子票
            收件人：测试收件人
            联系电话：13800000000
            寄送地址：测试省测试市测试路 1 号
            邮箱：mailto:invoice@example.com
            开票备注：对公打款
            """;

        var mail = new MailEnvelope(12, "m-1", "sino-esign@sinotrans.com", "中外运向您提交了开票申请", DateTimeOffset.Now, body);
        var result = new InvoiceParser().Parse(mail, "test@example.com");

        Assert.True(result.Success);
        Assert.Equal("测试物流有限公司", result.Application!.CompanyName);
        Assert.Equal("TEST-CREDIT-20260808", result.Application.CreditCode);
        Assert.Equal(300m, result.Application.Amount);
        Assert.Equal(new DateTime(2026, 8, 7, 10, 28, 0), result.Application.ApplyTime);
        Assert.Equal("invoice@example.com", result.Application.Email);
        Assert.Equal("对公打款", result.Application.Remark);
    }

    [Fact]
    public void AcceptsEnglishColonAndHtmlBreaks()
    {
        const string body = "公司名称: 测试物流<br>信用代码: 91370000TEST<br>申请金额: 12.50 元<br>申请时间: 2026-08-08 09:30<br>邮箱: finance@example.com";
        var mail = new MailEnvelope(13, "m-2", "sino-esign@sinotrans.com", "中外运向您提交了开票申请", DateTimeOffset.Now, body);
        var result = new InvoiceParser().Parse(mail, "test@example.com");

        Assert.True(result.Success);
        Assert.Equal(12.50m, result.Application!.Amount);
        Assert.Equal("finance@example.com", result.Application.Email);
    }

    [Fact]
    public void RejectsMissingRequiredField()
    {
        var mail = new MailEnvelope(14, "m-3", "sino-esign@sinotrans.com", "中外运向您提交了开票申请", DateTimeOffset.Now,
            "公司名称：测试企业\n申请金额：10 元\n申请时间：2026-08-07 10:28\n邮箱：a@example.com");
        var result = new InvoiceParser().Parse(mail, "test@example.com");

        Assert.False(result.Success);
        Assert.Contains("信用代码", result.MissingFields);
    }

    [Fact]
    public void FallbackHashDistinguishesDifferentApplications()
    {
        var first = new InvoiceApplication { CreditCode = "A", ApplyTime = new DateTime(2026, 8, 8, 10, 0, 0), Amount = 100m, MailFrom = "sino-esign@sinotrans.com" };
        var second = new InvoiceApplication { CreditCode = "A", ApplyTime = new DateTime(2026, 8, 8, 10, 1, 0), Amount = 100m, MailFrom = "sino-esign@sinotrans.com" };
        Assert.NotEqual(Deduplication.CreateFallbackHash(first), Deduplication.CreateFallbackHash(second));
    }

    [Fact]
    public void ParsesHtmlParagraphsCurrencyAndMailto()
    {
        const string body = "<p>邮箱：<a href=\"mailto:finance@example.com\">finance@example.com</a></p>" +
            "<div>申请时间：2026-08-08 09:30</div><div>申请金额：￥1,300.00 元</div>" +
            "<p>信用代码：000123</p><p>公司名称：测试物流</p>";
        var mail = new MailEnvelope(15, "m-4", "sino-esign@sinotrans.com", "中外运向您提交了开票申请", DateTimeOffset.Now, body);

        var result = new InvoiceParser().Parse(mail, "test@example.com");

        Assert.True(result.Success);
        Assert.Equal(1300m, result.Application!.Amount);
        Assert.Equal("000123", result.Application.CreditCode);
        Assert.Equal("finance@example.com", result.Application.Email);
    }

    [Fact]
    public void RejectsInvalidEmailAddress()
    {
        var mail = new MailEnvelope(16, "m-5", "sino-esign@sinotrans.com", "中外运向您提交了开票申请", DateTimeOffset.Now,
            "公司名称：测试企业\n信用代码：CODE\n申请金额：10\n申请时间：2026-08-08 09:30\n邮箱：不是邮箱");

        var result = new InvoiceParser().Parse(mail, "test@example.com");

        Assert.False(result.Success);
        Assert.Contains("邮箱", result.MissingFields);
    }
}

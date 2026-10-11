using InvoiceMailAssistant.App;
using Xunit;

namespace InvoiceMailAssistant.Tests;

public sealed class InvoiceSourceTests
{
    [Fact]
    public void ParsesYuntianhuaSampleShapeAndPreservesSource()
    {
        const string body = """
            云天化电子签服务平台向您提交了开票申请，申请信息如下：
            公司名称：测试新材料有限公司
            信用代码：91320000TESTSOURCE01
            申请金额：9.52 元
            申请时间：2026-10-09 15:52
            开票方式：电子票
            收件人：测试收件人
            联系电话：13800000000
            寄送地址：测试市测试区测试路 2 号
            邮箱：invoice@example.com
            """;

        var mail = new MailEnvelope(
            901,
            "<yuntianhua-source@example.test>",
            "cs@qsign.cn",
            "电子签服务平台开票申请",
            new DateTimeOffset(2026, 10, 9, 7, 52, 0, TimeSpan.Zero),
            body);

        var result = new InvoiceParser().Parse(mail, "test-yuntianhua@example.com");

        Assert.True(result.Success);
        Assert.Equal(InvoiceSource.Yuntianhua, result.Application!.Source);
        Assert.Equal("云天化", result.Application.DisplaySourceName);
        Assert.Equal(9.52m, result.Application.Amount);
    }

    [Fact]
    public void SinotransRemainsTheDefaultKnownSourceForLegacyEnvelopeShape()
    {
        var mail = new MailEnvelope(
            902,
            "<sinotrans-source@example.test>",
            "sino-esign@sinotrans.com",
            "中外运向您提交了开票申请",
            DateTimeOffset.UtcNow,
            "公司名称：测试企业\n信用代码：CODE\n申请金额：10 元\n申请时间：2026-10-09 15:52\n邮箱：invoice@example.com");

        var result = new InvoiceParser().Parse(mail, "account@example.com");

        Assert.True(result.Success);
        Assert.Equal(InvoiceSource.Sinotrans, result.Application!.Source);
    }

    [Fact]
    public void RejectsInvoiceBodyFromAnUnknownSender()
    {
        var mail = new MailEnvelope(
            903,
            "<unknown-source@example.test>",
            "spoof@example.com",
            "开票申请",
            DateTimeOffset.UtcNow,
            "云天化电子签服务平台向您提交了开票申请\n公司名称：测试企业");

        Assert.False(InvoiceSourceCatalog.TryClassify(mail, out _));
    }

    [Fact]
    public void FormatsNotificationsWithEachSourceName()
    {
        var text = SourceNotificationFormatter.Format(
        [
            new InvoiceSourceCount(InvoiceSource.Sinotrans, "中外运", 2),
            new InvoiceSourceCount(InvoiceSource.Yuntianhua, "云天化", 1)
        ]);

        Assert.Contains("来自中外运 2 条", text);
        Assert.Contains("来自云天化 1 条", text);
        Assert.Contains("已写入或进入等待队列", text);
    }
}

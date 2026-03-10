# Auto Forward As New
## 场景

**仅适用于个人 Gmail 邮箱以及任何利用 Gmail 搭建的企业、学校等的工作邮箱。**

当使用传统的邮件列表（如 Mailman 2）发送邮件，并通过自动转发至其他邮箱（如Outlook/MSN/Hotmail）时，由于邮件列表修改了邮件主题或正文，破坏了原始的 DKIM 签名与 ARC 信任链。对于采用严格的 DMARC等验证策略的邮件服务商可能会将当次转发判定为伪造或诈骗邮件并直接拒收，产生 `5322.From` 报错。

这个问题虽然可以通过使用支持 ARC 的邮件列表等方法彻底解决，但是由于说服你所属的设施更换或更新软件通常是困难的，因此在个人端解决这个问题是更可行的替代解决方法。

## 功能

通过部署在 [Google Apps Script](https://script.google.com/) 自动化脚本监听退信错误通知，从当前地址名义重新发送被退回的邮件，解决所有 SPF/DKIM 及其他验证策略方面的问题。

Example: [NDR error "550 5.7.515" in Outlook.com](https://support.microsoft.com/en-us/topic/fix-ndr-error-550-5-7-515-in-outlook-com-34cfe8f8-6fbf-457e-9e8b-9e4dbaf4e0ef)

```
550 5.7.515 Access denied, sending domain <domain> does not meet the required authentication level.
The sender's domain in the 5322.From address doesn't meet the authentication requirements defined for the sender.
```

## 部署

- 登录需要转发的 Gmail 账号，访问 [Google Apps Script](https://script.google.com/)；

- 新建项目，并将“appsScript.gs”的全部内容覆盖默认占位代码；

- 在代码顶部，修改以下全局常量为您自己的目标接收邮箱：

```
const targetEmail = "change+into+your+email+address@example.com"; // 替换为转发至的邮箱地址
```

- 保存项目，选择要运行的函数为“main”，点击运行并授权当前项目访问 Gmail；
- 为项目添加触发器，“选择摇匀行的功能”选择“main”，“活动来源”选择“时间驱动”，同时自行调整触发周期（默认为每小时一次）；
- 保存触发器后，脚本自动在云端运行。

## 注意

- 部署时应当确保正在使用 Gmail 的“自动转发”功能，脚本只会在接收到投递失败通知时才会尝试查找并重发出错邮件；

```
const searchQuery = "is:unread from:mailer-daemon@googlemail.com subject:Delivery Status Notification (Failure)";
```

- 配合使用 Gmail 自带的过滤器，可以自动将“mailer-daemon@googlemail.com”的错误通知归档，不影响本脚本运行，同时无感解决此类 NDR 错误。
- 但是注意务必不要在过滤时直接将错误通知邮件设置已读，也不要主动将通知已读，否则脚本会误认为当次报错已经处理而错误跳过重发。














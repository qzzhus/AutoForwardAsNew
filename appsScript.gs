//========================================================
// 自动化脚本：利用报错反馈自动查找退回邮件，从当前地址作为新邮件发送
// 作者：qzzhus
// 日期：2026-03-09
// 更新日期：2026-05-19
//========================================================
// 配置
//--------------------------------------------------------
// 转发的目标邮箱地址
const targetEmail = "change+into+your+email+address@example.com";
// 搜索 NDR 错误通知邮件的关键词
const searchQuery = "is:unread from:mailer-daemon@googlemail.com subject:Delivery Status Notification (Failure)";
// 标记被处理的原邮件
const FORWARDED_LABEL_NAME = "已转发";
//========================================================

// 脚本中选择运行此函数
function main(){
  var threads = GmailApp.search(searchQuery, 0, 10);
  if (!threads.length) { 
  console.info("【无行动】没有未处理的错误报告");
  return;
  }

  var processedLabel = GmailApp.getUserLabelByName(FORWARDED_LABEL_NAME);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(FORWARDED_LABEL_NAME);
  }

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = messages.length - 1; j >= 0; j--) {
      var errorMsg = messages[j];
      if (!errorMsg.isUnread()) continue;
      findMail(errorMsg, processedLabel);
    }
  }
}

// 查找匹配错误邮件
function findMail(errorMsg, processedLabel) {
  var targetMsg = null;
  var originalMsgId = null;

  // ==========================================
  // 默认方法：尝试提取 Message-ID 进行精确匹配（2026-05-19）
  // ==========================================
  // 优先尝试：解包附件提取 (针对标准 RFC822 退信格式，含 noname 等附件)
  var attachments = errorMsg.getAttachments();
  for (var a = 0; a < attachments.length; a++) {
    try {
      var attContent = attachments[a].getDataAsString();
      var attMatch = attContent.match(/^Message-ID:\s*<([^>]+)>/im);
      if (attMatch && attMatch[1]) {
        originalMsgId = attMatch[1].trim();
        console.info("    【精确匹配】Message-ID = " + originalMsgId + " 来自 " + attachments[a].getName());
        break;
      }
    } catch (e) {
      // 忽略无法被转换为纯文本的二进制附件（如图片等）
      console.warn("    【精确匹配】跳过非文本附件: " + attachments[a].getName());
    }
  }
  // 回退尝试：如果附件里没有，再查找原始邮件
  if (!originalMsgId) {
    var rawContent = errorMsg.getRawContent();
    var msgIdMatches = rawContent.match(/^Message-ID:\s*<([^>]+)>/gmi);
    if (msgIdMatches && msgIdMatches.length > 1) {
      // 获取源码中最后一个 Message-ID
      var lastMatch = msgIdMatches[msgIdMatches.length - 1];
      originalMsgId = lastMatch.replace(/^Message-ID:\s*</i, "").replace(/>$/, "").trim();
      console.info("    【精确匹配】Message-ID = " + originalMsgId);
    }
  }
  // 从 Message-ID 精确搜索
  if (originalMsgId) {
    var exactThreads = GmailApp.search("rfc822msgid:" + originalMsgId);
    if (exactThreads.length > 0) {
      var exactMsgs = exactThreads[0].getMessages();
      for (var k = 0; k < exactMsgs.length; k++) {
        if (exactMsgs[k].getHeader("Message-ID").indexOf(originalMsgId) !== -1) {
          targetMsg = exactMsgs[k];
          console.info("    【精确匹配】成功！");
          break;
        }
      }
    }
  }

  // ==========================================
  // 回退方法：如果精确匹配失败，降级使用时间模糊比对
  // ==========================================
  if (!targetMsg) {
    console.info("    【模糊匹配】从 Message-ID 匹配失败，回退至时间匹配");
    // 构建搜索查询：查找报错发生前 10 分钟至发生后 1 分钟收到的，且非退信的邮件
    var errorTime = errorMsg.getDate().getTime();
    var epochSec = Math.floor(errorTime / 1000); // 转换为搜索用的秒级时间戳
    var origQuery = "before:" + (epochSec + 60) + " after:" + (epochSec - 600) + " -from:mailer-daemon@googlemail.com";
    var origThreads = GmailApp.search(origQuery, 0, 10);
    // 遍历候选邮件，寻找时间差最小的原始邮件
    var minDiff = Infinity;
    for (var t = 0; t < origThreads.length; t++) {
      var origMsgs = origThreads[t].getMessages();
      for (var m = 0; m < origMsgs.length; m++) {
        var candidateMsg = origMsgs[m];
        var candidateTime = candidateMsg.getDate().getTime();
        var diff = errorTime - candidateTime;
        if (diff > 0 && diff < minDiff) {
          minDiff = diff;
          targetMsg = candidateMsg;
        }
      }
    }
  }

  // ==========================================
  // 转发目标邮件并添加标签；报错邮件标记已读并移动
  // ==========================================
  if (targetMsg) {
    fwdAsNew(targetMsg);
    targetMsg.getThread().addLabel(processedLabel);
    errorMsg.markRead();
    errorMsg.getThread().moveToArchive();
    return;
  }
  console.error("【发送失败】没有匹配到退回邮件")
}

function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/^"|"(?=\s*<)/g, "")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function fwdAsNew(msg){
  // 原邮件内容
  var originalSubject = msg.getSubject();
  var originalSender = msg.getFrom();
  var originalDate = msg.getDate();
  var originalBody = msg.getBody(); 
  var attachments = msg.getAttachments(); 
  var originalTo = msg.getTo();
  var originalCc = msg.getCc();

  // 转义时间，"Thursday, January 1, 2026 1:00:00 PM"形式
  var timeZone = Session.getScriptTimeZone();
  var formattedDate = Utilities.formatDate(originalDate, timeZone, "EEEE, MMMM d, yyyy h:mm:ss a");
  
  // 转义收发件人，"sender's name <address@example.com>"形式
  var safeSender = escapeHtml(originalSender);
  var safeTo = escapeHtml(originalTo);
  var safeCc = escapeHtml(originalCc);

  // 提取发件人名
  var cleanSenderName = originalSender.replace(/^"|"(?=\s*<)/g, "").replace(/<.*>/, "").trim();
  if (!cleanSenderName) cleanSenderName = originalSender;

  // 新邮件头部
  var headerHtml = 
    "<div style='font-family: Calibri, Arial, sans-serif; font-size: small;'>" +
    "<b>发件人:</b> " + safeSender + "<br>" +
    "<b>发送时间:</b> " + formattedDate + "<br>" +
    "<b>主题:</b> " + originalSubject + "<br>" +
    "<b>收件人:</b> " + safeTo + "<br>";
  if (safeCc) {
    headerHtml += "<b>抄送:</b> " + safeCc + "<br>";
  }
  headerHtml += "</div><hr style=\"display:inline-block;width:98%\" class=\"\"><br><br>";

  // 新邮件内容
  // 格式化纯文本邮件（2026-04-06）
  var isHtmlEmail = /<(br|div|p|html|body|table|span)/i.test(originalBody);
  var formattedBody = "";
  if (!isHtmlEmail) {
    formattedBody = msg.getPlainBody().replace(/\r?\n/g, '<br>');
  } else {
    formattedBody = originalBody;
  }
  var newBody = headerHtml + formattedBody;
  
  // 作为新邮件发送
  // 回复至原发件人
  // 新邮件发件人名为"[FwdOnERR] original sender's name"形式
  GmailApp.sendEmail("", originalSubject, "", {
    htmlBody: newBody,
    bcc: targetEmail,
    attachments: attachments,
    replyTo: originalSender, 
    name: "[FwdOnERR] " + cleanSenderName
  });

  // 控制台通知
  console.info(`【发送成功】
主题：${originalSubject}
发件人：${cleanSenderName}
原始发送时间：${formattedDate}
样式：${isHtmlEmail ? "html" : "plain"}`);
}




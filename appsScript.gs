//========================================================
// 自动化脚本：利用报错反馈自动查找退回邮件，从当前地址作为新邮件发送
// 作者：qzzhus
// 日期：2026-03-09
// 更新日期：2026-04-06
//========================================================

// 替换为转发至的邮箱地址
const targetEmail = "change+into+your+email+address@example.com";
// 搜索 NDR 错误通知邮件的关键词，某些情况可能需要自行修改
const searchQuery = "is:unread from:mailer-daemon@googlemail.com subject:Delivery Status Notification (Failure)";
const lang = "zh-cn";

function main(){
  var threads = GmailApp.search(searchQuery, 0, 10);
  if (!threads.length) { 
  console.info("【无行动】没有未处理的错误报告");
  return;
  }
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var errorMsg = messages[j];
      if (!errorMsg.isUnread()) continue;
      findMail(errorMsg);
    }
  }
}

function findMail(errorMsg) {
  var errorTime = errorMsg.getDate().getTime();
  
  // 查找报错发生前 10 分钟至发生后 1 分钟收到的，且非退信的邮件
  var epochSec = Math.floor(errorTime / 1000);
  var origQuery = "before:" + (epochSec + 60) + " after:" + (epochSec - 600) + " -from:mailer-daemon@googlemail.com";
  var origThreads = GmailApp.search(origQuery, 0, 10);

  // 遍历候选邮件，寻找时间差最小的原始邮件
  var targetMsg = null;
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

  // 转发目标邮件，将报错邮件标记已读并移动
  if (targetMsg) {
    fwdAsNew(targetMsg);
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
  // if (!msg.isUnread()) return;

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




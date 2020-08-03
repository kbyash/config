#!/usr/bin/osascript

function run(workTime) {
    var outlook = Application("Microsoft Outlook")
    var date = new Date()
    var year = date.getFullYear()
    var dayAndMonth = toDoubleDigits(date.getMonth() + 1) + "/" + toDoubleDigits(date.getDate()) // getMonth(): 0 ~ 11 のため1を足している
    var startHour = "10:00"
    var minute = toDoubleDigits(date.getMinutes())
    var endHour = "19:00"//toDoubleDigits(date.getHours() + parseInt(workTime.length ? workTime : 9))
    var workType = (workTime.length ? (workTime + "h") : "終日") + "WFH" // {workTime}hWFH or 終日WFH

    var subject = "kbyash@ " + dayAndMonth + " " + workType
    var content = "<font face='游ゴシック'>"
        + "本日は" + workType + "で勤務いたします。<br>"
        + "WFH開始: " + year + "/" + dayAndMonth + " " + startHour  + "<br>"
        + "WFH終了: " + year + "/" + dayAndMonth + " " + endHour  + "（予定）<br>"
        + "休憩時間: 1時間 (予定)<br>"
        + "よろしくお願いします。<br>"
        + "</font>"

    var message = outlook.OutgoingMessage({ subject: subject, content: content }).make()
    message.toRecipients.push(outlook.ToRecipient({ emailAddress: { address: "aws-dev-support-jp-ps-kintai@amazon.com" } }))
    outlook.send(message) // open(message) を send(message) にすると送信できます
}

// 日付や時間の0 詰め
function toDoubleDigits(num) {
    return ("0" + num).slice(-2)
}

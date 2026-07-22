sub init()
    m.dateTimeLabel = m.top.findNode("dateTimeLabel")
    m.label = m.top.findNode("timerLabel")
    m.subtitle = m.top.findNode("subtitleLabel")
    m.timer = m.top.findNode("countdownTimer")
    m.syncTimer = m.top.findNode("syncTimer")
    m.clockTimer = m.top.findNode("clockTimer")
    m.top.screensaverTimeout = 0
    m.top.enableScreenSaver = false
    m.remainingMillis = 0
    m.hasLoaded = false
    m.timerStarted = false
    m.requestInFlight = false
    m.serverRunning = false
    m.apiConnected = false
    m.dateTimeLabel.font.size = 39
    m.label.font.size = 168
    m.subtitle.font.size = 48
    m.label.text = "--:--"
    m.subtitle.text = "CARGANDO TIEMPO"

    updateDateTime()
    m.timer.observeField("fire", "updateCountdown")
    m.syncTimer.observeField("fire", "fetchTimer")
    m.clockTimer.observeField("fire", "updateDateTime")
    m.clockTimer.control = "start"
    m.syncTimer.control = "start"
    fetchTimer()
end sub

function twoDigits(value as Integer) as String
    text = value.ToStr()
    if value < 10 then
        text = "0" + text
    end if

    return text
end function

sub updateDateTime()
    now = CreateObject("roDateTime")
    now.ToLocalTime()

    dayText = twoDigits(now.GetDayOfMonth())
    monthText = twoDigits(now.GetMonth())
    yearText = now.GetYear().ToStr()
    hourText = twoDigits(now.GetHours())
    minuteText = twoDigits(now.GetMinutes())
    secondText = twoDigits(now.GetSeconds())

    m.dateTimeLabel.text = dayText + "/" + monthText + "/" + yearText + " " + hourText + ":" + minuteText + ":" + secondText
end sub

sub fetchTimer()
    if m.requestInFlight then
        return
    end if

    m.requestInFlight = true
    m.task = CreateObject("roSGNode", "ApiTask")
    m.task.observeField("response", "onResponse")
    m.task.observeField("error", "onApiError")
    m.task.control = "RUN"
end sub

sub onApiError()
    m.requestInFlight = false

    m.subtitle.text = "ERROR: " + m.task.error
    m.apiConnected = false

    if m.timerStarted then
        stopCountdown()
    end if

    if not m.hasLoaded then
        m.label.text = "--:--"
        m.label.visible = true
    end if
end sub

sub onResponse()
    m.requestInFlight = false
    data = ParseJson(m.task.response)

    if data = invalid then
        if not m.hasLoaded then
            m.label.text = "--:--"
            m.subtitle.text = "NO SE PUDO LEER EL TIEMPO"
        end if
        return
    end if

    if data.remaining = invalid then
        if not m.hasLoaded then
            m.label.text = "--:--"
            m.subtitle.text = "NO SE PUDO LEER EL TIEMPO"
        end if
        return
    end if

    serverRemaining = data.remaining
    serverRunning = true
    if data.running <> invalid then
        serverRunning = data.running
    end if

    diff = serverRemaining - m.remainingMillis
    if diff < 0 then
        diff = -diff
    end if

    if not m.hasLoaded or diff > 3000 or serverRunning <> m.serverRunning then
        m.remainingMillis = serverRemaining
        renderCountdown()
    end if

    m.serverRunning = serverRunning
    m.hasLoaded = true
    m.apiConnected = true
    m.subtitle.text = "TIEMPO RESTANTE"

    if m.serverRunning and m.remainingMillis > 0 then
        startCountdown()
    else
        stopCountdown()
    end if
end sub

sub startCountdown()
    if m.timer = invalid then
        m.label.text = "TIMER INVALID"
        return
    end if

    if m.timerStarted then
        return
    end if

    m.timer.control = "start"
    m.timerStarted = true
end sub

sub stopCountdown()
    if m.timer = invalid then
        return
    end if

    m.timer.control = "stop"
    m.timerStarted = false
end sub

sub updateCountdown()
    m.remainingMillis = m.remainingMillis - 1000

    if m.remainingMillis <= 0 then
        m.remainingMillis = 0
        renderCountdown()
        m.timer.control = "stop"
        m.timerStarted = false
        return
    end if

    renderCountdown()
end sub

sub renderCountdown()
    totalSeconds = Int(m.remainingMillis / 1000)

    hours = Int(totalSeconds / 3600)
    minutes = Int((totalSeconds mod 3600) / 60)
    seconds = totalSeconds Mod 60

    minuteText = minutes.ToStr()
    if minutes < 10 then
        minuteText = "0" + minuteText
    end if

    secondText = seconds.ToStr()
    if seconds < 10 then
        secondText = "0" + secondText
    end if

    if hours > 0 then
        m.label.text = hours.ToStr() + ":" + minuteText + ":" + secondText
    else
        m.label.text = minuteText + ":" + secondText
    end if

    if totalSeconds <= 60 then
        m.label.color = "0xFF0000FF"
    else if totalSeconds <= 300 then
        m.label.color = "0xFFFF00FF"
    else
        m.label.color = "0xFFFFFFFF"
    end if

    if totalSeconds > 0 and totalSeconds <= 10 then
        m.label.visible = ((totalSeconds mod 2) = 0)
    else
        m.label.visible = true
    end if
end sub

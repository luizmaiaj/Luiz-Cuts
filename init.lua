local hyper = { "cmd", "alt", "ctrl", "shift" }
local topLeft = {x=0, y=0, w=0.5, h=0.5}
local topRight = {x=0.5, y=0, w=0.5, h=0.5}
local bottomLeft = {x=0, y=0.5, w=0.5, h=0.5}
local bottomRight = {x=0.5, y=0.5, w=0.5, h=0.5}
local left = {x=0, y=0, w=0.5, h=1}
local right = {x=0.5, y=0, w=0.5, h=1}

-- RELOAD CONFIG
hs.hotkey.bind(hyper, "f12", function()
  hs.console.clearConsole()
  hs.reload()
end)
hs.alert.show("Config loaded")

-- Set a timer to capture active app and title every 30 seconds
-- {year = 1998, month = 9, day = 16, yday = 259, wday = 4, hour = 23, min = 48, sec = 10, isdst = false}
togglCapture = hs.timer.doEvery(30, function()

  day = os.date("*t")

  fileName = hs.fs.currentDir() .. "/" .. day.year .. string.format("%02d", day.month) .. string.format("%02d", day.day) .. ".csv"

  windowMouse = get_window_under_mouse()

  if windowMouse ~= nil then
    info = windowMouse:application():title() .. "; " .. windowMouse:title() .. "; " .. string.format("%02d", day.hour) .. ":" .. string.format("%02d", day.min) .. ":" .. string.format("%02d", day.sec) .. "\n"

    -- hs.alert.show(info)

    if not file_exists(fileName) then
      file = io.open(fileName, "w")

      io.output(file)

      io.write("mouse app, mouse title, time\n")

      io.close(file)
    end

    file = io.open(fileName, "a")

    io.output(file)

    io.write(info)

    io.close(file)
  end

end)

function file_exists(name)
  local f=io.open(name,"r")
  if f~=nil then io.close(f) return true else return false end
end

-- get window under mouse cursor
function get_window_under_mouse()
  -- Invoke `hs.application` because `hs.window.orderedWindows()` doesn't do it
  -- and breaks itself
  local _ = hs.application

  local my_pos = hs.geometry(hs.mouse.absolutePosition())
  local my_screen = hs.mouse.getCurrentScreen()

  return hs.fnutils.find(hs.window.orderedWindows(), function(w)
    return my_screen == w:screen() and my_pos:inside(w:frame())
  end)
end

-- move the window under mouse cursor
function move_window_under_mouse(p) -- p for position name

  local w = get_window_under_mouse()
  local wf = w:frame() -- get window frame
  local f = w:screen():frame()

  local moveTo = {right={x=math.floor(f.x+(f.w/2)), y=math.floor(f.y), w=math.floor(f.w/2), h=math.floor(f.h)},
  left={x=math.floor(f.x), y=math.floor(f.y), w=math.floor(f.w/2), h=math.floor(f.h)},
  center={x=math.floor(f.x+(f.w/6)), y=math.floor(f.y+(f.h/6)), w=math.floor((f.w/6)*4), h=math.floor((f.h/6)*4)},
  maximise={x=math.floor(f.x), y=math.floor(f.y), w=math.floor(f.w), h=math.floor(f.h)},
  top={x=math.floor(f.x), y=math.floor(f.y), w=math.floor(f.w), h=math.floor(f.h/2)},
  topLeft={x=math.floor(f.x), y=math.floor(f.y), w=math.floor(f.w/2), h=math.floor(f.h/2)},
  topRight={x=math.floor(f.x+(f.w/2)), y=math.floor(f.y), w=math.floor(f.w/2), h=math.floor(f.h/2)},
  bottom={x=math.floor(f.x), y=math.floor(f.y+(f.h/2)), w=math.floor(f.w), h=math.floor(f.h/2)},
  bottomLeft={x=math.floor(f.x), y=math.floor(f.y+(f.h/2)), w=math.floor(f.w/2), h=math.floor(f.h/2)},
  bottomRight={x=math.floor(f.x+(f.w/2)), y=math.floor(f.y+(f.h/2)), w=math.floor(f.w/2), h=math.floor(f.h/2)},
  leftTwoThirds={x=math.floor(f.x), y=math.floor(f.y), w=math.floor((f.w/3)*2), h=math.floor(f.h)},
  rightOneThird={x=math.floor(f.x+(f.w/3)*2), y=math.floor(f.y), w=math.floor(f.w/3), h=math.floor(f.h)},
  leftFourFifths={x=math.floor(f.x), y=math.floor(f.y), w=math.floor((f.w/5)*4), h=math.floor(f.h)},
  rightOneFifth={x=math.floor(f.x+(f.w/5)*4), y=math.floor(f.y), w=math.floor(f.w/5), h=math.floor(f.h)},
  appstore={x=112, y=112, w=(1280-224)/2, h=(800-224)/2}}

  --hs.console.printStyledtext("===============================")
  --hs.console.printStyledtext("pressed key", p)
  --hs.console.printStyledtext("window frame: ", wf)

  local rect = moveTo[p]
  --hs.console.printStyledtext("moveTo before: ", rect.x, rect.y, rect.w, rect.h)

  if wf == moveTo["left"] then -- if it's left already

    if p == "left" then
      p = "leftTwoThirds"
    elseif p == "top" then
      p = "topLeft"
    elseif p == "bottom" then
      p = "bottomLeft"
    end
    
  elseif wf == moveTo["leftTwoThirds"] then -- if it's left two thirds already
    p = "leftFourFifths"
  elseif wf == moveTo["right"] then -- if it's right already

    if p == "right" then
      p = "rightOneThird"
    elseif p == "top" then
      p = "topRight"
    elseif p == "bottom" then
      p = "bottomRight"
    end

  elseif wf == moveTo["rightOneThird"] then -- if it's top right already
    p = "rightOneFifth"
  elseif wf == moveTo["topRight"] then -- if it's top right already

    if p == "left" then
      p = "topLeft"
    elseif p == "bottom" then
      p = "bottomRight"
    end

  elseif wf == moveTo["topLeft"] then -- if it's top left already

    if p == "right" then
      p = "topRight"
    elseif p == "bottom" then
      p = "bottomLeft"
    end

  elseif wf == moveTo["bottomRight"] then -- if it's top right already

    if p == "top" then
      p = "topRight"
    elseif p == "left" then
      p = "bottomLeft"
    end

  elseif wf == moveTo["bottomLeft"] then -- if it's top left already

    if p == "top" then
      p = "topLeft"
    elseif p == "right" then
      p = "bottomRight"
    end

  end

  --hs.console.printStyledtext("converted key", p)
  local rect = moveTo[p]
  --hs.console.printStyledtext("moveTo after: ", rect.x, rect.y, rect.w, rect.h)

  w:focus()

  if p ~= "focus" then
    w:setFrame(moveTo[p], 0)

    -- move mouse to the center of the recently moved window
    local current_pos = hs.geometry(hs.mouse.absolutePosition()) -- get current mouse position
    local frame = w:frame() -- get updated window frame
    if not current_pos:inside(frame) then
      hs.mouse.absolutePosition(frame.center)
    end

    mouseHighlight()

  end

end

hs.hotkey.bind(hyper, "`", "focus", function() -- JUST FOCUS
  move_window_under_mouse("focus")
end)

hs.hotkey.bind(hyper, "-", "center", function() -- CENTER
  move_window_under_mouse("center")
end)

hs.hotkey.bind(hyper, "=", "maximise", function() -- MAXIMISE
  move_window_under_mouse("maximise")
end)

hs.hotkey.bind(hyper, "up", "up", function() -- UP
  move_window_under_mouse("top")
end)

hs.hotkey.bind(hyper, "down", "down", function() -- DOWN
  move_window_under_mouse("bottom")
end)

hs.hotkey.bind(hyper, "0", "appstore", function() -- APPSTORE SIZE
  move_window_under_mouse("appstore")
end)

hs.hotkey.bind(hyper, "left", "left", function() -- LEFT
  move_window_under_mouse("left")
end)

hs.hotkey.bind(hyper, "right", "right", function() -- RIGHT
  move_window_under_mouse("right")
end)

-- primary screen
hs.hotkey.bind(hyper, "1", function()
  local screens = hs.screen.allScreens()

  local w = get_window_under_mouse()

  if #(screens) > 1 then
    if screens[1] ~= w:screen() then

      hs.alert.show("primary screen")

      w:focus()

      w:moveToScreen(screens[1], false, true)
    end
  end
end)

-- secondary screen
hs.hotkey.bind(hyper, "2", function()
  local screens = hs.screen.allScreens()

  local w = get_window_under_mouse()

  if #(screens) > 1 then
    if screens[2] ~= w:screen() then

      hs.alert.show("secondary screen")

      w:focus()

      w:moveToScreen(screens[2], false, true)
    end
  end
end)

-- tertiary screen
hs.hotkey.bind(hyper, "3", function()
  local screens = hs.screen.allScreens()

  local w = get_window_under_mouse()

  if #(screens) > 2 then
    if screens[3] ~= w:screen() then

      hs.alert.show("tertiary screen")

      w:focus()

      w:moveToScreen(screens[3], false, true)
    end
  end
end)

-- HIGHLIGHT MOUSE
mouseCircle = nil
mouseCircleTimer = nil
function mouseHighlight()
    -- Delete an existing highlight if it exists
    if mouseCircle then
        mouseCircle:delete()
        if mouseCircleTimer then
            mouseCircleTimer:stop()
        end
    end
    -- Get the current co-ordinates of the mouse pointer
    mousepoint = hs.mouse.absolutePosition()
    -- Prepare a big red circle around the mouse pointer
    mouseCircle = hs.drawing.circle(hs.geometry.rect(mousepoint.x-40, mousepoint.y-40, 80, 80))
    mouseCircle:setStrokeColor({["red"]=1,["blue"]=0,["green"]=0,["alpha"]=1})
    mouseCircle:setFill(false)
    mouseCircle:setStrokeWidth(5)
    mouseCircle:show()

    -- Set a timer to delete the circle after 2 seconds
    mouseCircleTimer = hs.timer.doAfter(2, function()
      mouseCircle:delete()
      mouseCircle = nil
    end)
end
hs.hotkey.bind(hyper, "M", mouseHighlight)

-- WINDOWS CONFIG FOR WEEKLY REPORT
hs.hotkey.bind(hyper, "f1", "Weekly Report View", function()
  
  local screens = hs.screen.allScreens()

  local screenForMainApps = screens[1]
  local screenForSecApps = screens[1]

  if #(screens) > 1 then
    screenForMainApps = screens[2] --if more than one screen than main apps go on secondary screen
  end

  local windowLayout = {
      {"Microsoft Outlook", nil,    screenForSecApps,   left,       nil, nil},
      {"Slack",             nil,    screenForSecApps,   right,      nil, nil},
      {"Safari",            nil,    screenForMainApps,  bottomLeft, nil, nil},
      {"muCommander",       nil,    screenForMainApps,  topLeft,    nil, nil}
  }

  -- add weekly report window if found
  mo = hs.application.find("Microsoft Outlook")

  local moWindows = mo:allWindows()

  mwPri = hs.appfinder.windowFromWindowTitlePattern("PD Weekly Report.*")
  --mwPri = mo:findWindow("PD Weekly Report.*")
  if mwPri ~= nil then

    windowLayout[#(windowLayout)+1] = {"Microsoft Word",    mwPri,  screenForMainApps,  right,      nil, nil}

    mwPri:focus()
  end

  -- add temp file window if found
  mwSec = hs.appfinder.windowFromWindowTitlePattern("Temp.*")
  if mwSec ~= nil then

    windowLayout[#(windowLayout)+1] = {"Microsoft Word",    mwSec,  screenForMainApps,  topLeft,    nil, nil}

    mwSec:focus()
  end

  hs.alert.show(#(windowLayout))

  hs.layout.apply(windowLayout)

end)

-- WINDOWS CONFIG FOR MESSAGING
hs.hotkey.bind(hyper, "f2", "Messaging View", function()
  
  local screens = hs.screen.allScreens()

  local screenForMainApps = screens[1]
  local screenForSecApps = screens[1]

  if #(screens) > 1 then
    screenForMainApps = screens[2] --if more than one screen than main apps go on secondary screen
  end

  local windowLayout = {
      {"Microsoft Teams", nil, screenForMainApps, left,   nil, nil},
      {"Slack",             nil, screenForMainApps, right,  nil, nil},
      {"Safari",            nil, screenForSecApps,  left,   nil, nil},
      {"muCommander",       nil, screenForSecApps,  right,  nil, nil}
  }
  hs.layout.apply(windowLayout)

  hs.appfinder.appFromName("Microsoft Teams"):mainWindow():focus()
  hs.appfinder.appFromName("Slack"):mainWindow():focus()

end)

-- split top apps by order
hs.hotkey.bind(hyper, "F3", function()
  local windows = hs.window.orderedWindows()

  for i=1, 2 do
    local f = windows[i]:screen():frame()

    if i == 1 then
      windows[i]:setFrame({x=f.x+(f.w/2), y=f.y, w=f.w/2, h=f.h}, 0) -- right
    elseif i == 2 then
      windows[i]:setFrame({x=f.x, y=f.y, w=f.w/2, h=f.h}, 0) -- left
    end
  end

end)

SetCapsLockState, AlwaysOff

; Arrow keys like VIM
CapsLock & h::Send, {blind}{Left}
CapsLock & j::Send, {blind}{Down}
CapsLock & k::Send, {blind}{Up}
CapsLock & l::Send, {blind}{Right}
CapsLock & Space::Send, {blind}{Backspace}
CapsLock & $::Send, {blind}{End}
CapsLock & o::Send, {blind}{Home}
CapsLock & u::Send, {blind}{PgUp}
CapsLock & n::Send, {blind}{PgDn}

; Some special characters
CapsLock & z::Send, {blind}ß
CapsLock & /::Send, {blind}\


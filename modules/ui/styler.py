def change_colors(style):
    theme = {
        'disabledfg':"#eeeeee",
        'dark': "#777777",
        'darker': "#333333",
        'darkest': "#777777",
        'lighter': "#777777",
        'lightest': "#ffffff",
        'selectbg': "#41B1FF",
        'selectfg': "#ffffff",
        'foreground': "#111111",
        'background': "#dddddd",
        'borderwidth': 1,
        'font': ("Droid Sans", 10)
        }

    style.configure(".", padding=5, relief="flat", 
            background=theme['background'],
            foreground=theme['foreground'],
            bordercolor=theme['darker'],
            indicatorcolor=theme['selectbg'],
            focuscolor=theme['selectbg'],
            darkcolor=theme['dark'],
            lightcolor=theme['lighter'],
            troughcolor=theme['darker'],
            selectbackground=theme['selectbg'],
            selectforeground=theme['selectfg'],
            selectborderwidth=theme['borderwidth'],
            font=theme['font']
            )

    style.map(".",
        foreground=[('pressed', theme['darkest']), ('active', theme['selectfg'])],
        background=[('pressed', '!disabled', 'black'), ('active', theme['lighter'])]
        )

    style.configure("TButton", relief="flat")
    style.map("TButton", 
        background=[('disabled', theme['disabledfg']), ('pressed', theme['selectbg']), ('active', theme['selectbg'])],
        foreground=[('disabled', theme['disabledfg']), ('pressed', theme['selectfg']), ('active', theme['selectfg'])],
        bordercolor=[('alternate', theme['selectbg'])],
        )

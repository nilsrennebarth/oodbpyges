#!/usr/bin/env python3
"""
Prepare an interactive python - libreoffice session
"""

import uno
import types


class So(types.SimpleNamespace):
	pass


def init():
	localContext = uno.getComponentContext()
	resolver = localContext.ServiceManager.createInstanceWithContext(
		"com.sun.star.bridge.UnoUrlResolver",
		localContext
	)
	ctx = resolver.resolve(
		"uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
	)
	smgr = ctx.ServiceManager
	desktop = smgr.createInstanceWithContext(
		"com.sun.star.frame.Desktop", ctx
	)
	model = desktop.getCurrentComponent()
	return So(ctx=ctx, smgr=smgr, desktop=desktop, model=model)

; WebDriver Element class for Rufaydium
; By Xeo786

Class WDElement extends Session
{
	__new(Address,Element)
	{
		This.Address := Address
		This.Element := Element
	}
	
	TagName
	{
		get
		{
			return this.Send("name","GET")
		}
	}
	
	Rect()
	{
		return this.Send("rect","GET")
	}
	
	Size()
	{
		return this.Send("Size","GET")
	}
	
	Location()
	{
		return this.Send("location","GET")
	}
	
	LocationInView()
	{
		return this.Send("location_in_view","GET")
	}
	
	enabled()
	{
		return this.Send("enabled","GET")
	}
	
	Selected()
	{
		return this.Send("selected","GET")
	}
	
	Displayed()
	{
		return this.Send("displayed","POST",{"":""})
	}
	
	submit()
	{
		return this.Send("submit","POST",{"":""})
	}
	
	SendKey(text)
	{
		return this.Send("value","POST", {"text":text})
	}
	
	click()
	{
		return this.Send("click","POST",{"":""})
	}
	
	Move()
	{
		return this.Send("moveto","POST",{"element_id":this.id})
	}
	
	Title
	{
		get
		{
			return this.GetAttribute("title")
		}

		set
		{
			this.Execute("arguments[0].title = '" Value "'")
		}
	}

	Class
	{
		get
		{
			return this.GetAttribute("class")
		}

		set
		{
			this.Execute("arguments[0].className = '" Value "'")
		}
	}

	Name
	{
		get
		{
			return this.GetAttribute("name")
		}

		set
		{
			this.Execute("arguments[0].name = '" Value "'")
		}
	}

	id
	{
		get
		{
			return this.GetAttribute("id")
		}

		set
		{
			this.Execute("arguments[0].id = '" Value "'")
		}
	}

	value
	{
		get
		{
			v := this.Send("value","GET")
			if v.error
				return this.GetProperty("value")
			else
				return v	
		}
		
		Set
		{
			this.Clear()
			return this.Send("value","POST", {"text":Value})
		}
	}
	
	InnerText
	{
		get
		{
			e := this.Send("text","GET")
			if !e
				e := this.Execute("return arguments[0].InnerText")
			return e
		}

		set
		{
			this.Execute("arguments[0].innerText = '" Value "'")
		}
	}
	
	innerHTML
	{
		get
		{
			return  this.GetProperty("innerHTML")
		}

		set
		{
			this.Execute("arguments[0].innerHTML = '" Value "'")
		}
	}

	outerHTML
	{
		get
		{
			return  this.GetProperty("outerHTML")
		}

		set
		{
			this.Execute("arguments[0].outerHTML = '" Value "'")
		}
	}

	Clear()
	{
		;this.Send("ClearValue","POST"); not working for me
		obj :=  {"text": key.ctrl "a" key.delete}
		return this.Send("value","POST", obj)
	}
	
	GetAttribute(Name)
	{
		return this.Send("attribute/" Name,"GET")
	}

	GetProperty(Name)
	{
		return this.Send("property/" Name,"GET")
	}
	
	GetCSS(Name)
	{
		return this.Send("css/" Name,"GET")
	}
	
	ComputedRole() ; https://www.w3.org/TR/wai-aria-1.1/#usage_intro
	{
		return this.Send("computedrole","GET")
	}
	
	ComputedLable() ; https://www.w3.org/TR/wai-aria-1.1/#usage_intro
	{
		return this.Send("computedlabel","GET")
	}
	
	Uploadfile(filelocation)
	{
		return this.Send("file","POST",{})
	}
	
	Execute(script)
	{
		Origin := this.Address
		RegExMatch(Origin,"(.*)\/element\/(.*)$",i)
		args := [{This.Element:i2}]
		this.address := i1
		r := this.Send("execute/sync","POST", { "script":Script,"args":Args},1)
		this.address := Origin
		return r
	}

}

Class ShadowElement extends Session
{
	__new(Address)
	{
		This.Address := Address
	}
}

Class Key
{
	static Unidentified := "\uE000"
	static Cancel:= "\uE001"
	static Help:= "\uE002"
	static Backspace:= "\uE003"
	static Tab:= "\uE004"
	static Clear:= "\uE005"
	static Return:= "\uE006"
	static Enter:= "\uE007"
	static Shift:= "\uE008"
	static Control:= "\uE009"
	static Ctrl:= "\uE009"
	static Alt:= "\uE00A"
	static Pause:= "\uE00B"
	static Escape:= "\uE00C"
	static Space:= "\uE00D"
	static PageUp:= "\uE00E"
	static PageDown:= "\uE00F"
	static End:= "\uE010"
	static Home:= "\uE011"
	static ArrowLeft:= "\uE012"
	static ArrowUp:= "\uE013"
	static ArrowRight:= "\uE014"
	static ArrowDown:= "\uE015"
	static Insert:= "\uE016"
	static Delete:= "\uE017"
	static F1:= "\uE031"
	static F2:= "\uE032"
	static F3:= "\uE033"
	static F4:= "\uE034"
	static F5:= "\uE035"
	static F6:= "\uE036"
	static F7:= "\uE037"
	static F8:= "\uE038"
	static F9:= "\uE039"
	static F10:= "\uE03A"
	static F11:= "\uE03B"
	static F12:= "\uE03C"
	static Meta:= "\uE03D"
	static ZenkakuHankaku:= "\uE040"	
}

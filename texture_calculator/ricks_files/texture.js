function setSilt(sand, clay)
{
	document.info.silt.value = 100-parseFloat(sand)-parseFloat(clay);
}


function checkForm()
{
var sandtot;
var detailsc;

	if((document.info.vcs.value != "" && parseFloat(document.info.vcs.value) != 0) || (document.info.cs.value != "" && parseFloat(document.info.cs.value) != 0) || (document.info.ms.value != "" && parseFloat(document.info.ms.value) != 0) || (document.info.fs.value != "" && parseFloat(document.info.fs.value) != 0) || (document.info.vfs.value != "" && parseFloat(document.info.vfs.value) != 0))
	{
		detailsc = true;

		if(isNaN(parseFloat(document.info.vcs.value)))
		{
			document.info.vcs.value = 0;
			vcsFloat = 0;
		}
		else
		{
			vcsFloat = parseFloat(document.info.vcs.value);
		}

		if(isNaN(parseFloat(document.info.cs.value)))
		{
			document.info.cs.value = 0;
			csFloat = 0;
		}
		else
		{
			csFloat = parseFloat(document.info.cs.value);
		}

		if(isNaN(parseFloat(document.info.ms.value)))
		{
			document.info.ms.value = 0;
			msFloat = 0;
		}
		else
		{
			msFloat = parseFloat(document.info.ms.value);
		}

		if(isNaN(parseFloat(document.info.fs.value)))
		{
			document.info.fs.value = 0;
			fsFloat = 0;
		}
		else
		{
			fsFloat = parseFloat(document.info.fs.value);
		}

		if(isNaN(parseFloat(document.info.vfs.value)))
		{
			document.info.vfs.value = 0;
			vfsFloat = 0;
		}
		else
		{
			vfsFloat = parseFloat(document.info.vfs.value);
		}

		sandtot = (vcsFloat + csFloat + msFloat + fsFloat + vfsFloat).toPrecision(5)
		if( sandtot != parseFloat(document.info.sand.value).toPrecision(5))
		{
			alert('If you choose to enter values for sand components (' + vcsFloat + '+' + csFloat + '+' + msFloat + '+' + fsFloat + '+' + vfsFloat + '), they must add up to the value of sand - (' + document.info.sand.value + ')');
			return;
		}
	}
	else
	{
		detailsc = "false";
	}

	if(document.info.sand.value == "" || isNaN(parseFloat(document.info.sand.value)))
	{
		alert('You must insert a number for Percent Sand');
		document.info.sand.value = "";		
		return false;
	}
	else if(document.info.clay.value == "" || isNaN(parseFloat(document.info.clay.value)))
	{
		alert('You must insert a number for Percent Clay');
		document.info.clay.value = "";
		return false;
	}
	else if(parseFloat(document.info.sand.value) + parseFloat(document.info.clay.value) > 100)
	{
		alert('The total of Percent Sand and Percent Clay must be less than or equal to 100');
		return false;
	}

	else
	{
		document.info.sand.value = parseFloat(document.info.sand.value);
		document.info.clay.value = parseFloat(document.info.clay.value);
		setSilt(document.info.sand.value, document.info.clay.value);
		getTexture(document.info.sand.value, document.info.clay.value, document.info.silt.value, detailsc);
		drawPoint(document.info.sand.value, document.info.clay.value);
		return true;
	}
}


function getTexture(sand, clay, silt, detailsc)
{
	sand = parseFloat(sand);
	clay = parseFloat(clay);
	silt = parseFloat(silt);
	
	if((silt + 1.5*clay) < 15)
	{
		
		if(detailsc)
		{
			getSand(document.info.vcs.value, document.info.cs.value, document.info.ms.value, document.info.fs.value, document.info.vfs.value);
		}
		else
		{			
			document.info.texture.value = 'Sand';
		}
	}
	else if((silt + 1.5*clay >= 15) && (silt + 2*clay < 30))
	{
		if(detailsc)
		{
			getLoamSand(document.info.vcs.value, document.info.cs.value, document.info.ms.value, document.info.fs.value, document.info.vfs.value);
		}
		else
		{
			document.info.texture.value = 'Loamy Sand';
		}
	}
	else if((clay >= 7 && clay < 20) && (sand > 52) && ((silt + 2*clay) >= 30) || (clay < 7 && silt < 50 && (silt+2*clay)>=30))
	{
		if(detailsc)
		{
			getSandLoam(document.info.vcs.value, document.info.cs.value, document.info.ms.value, document.info.fs.value, document.info.vfs.value);
		}
		else
		{
			document.info.texture.value = 'Sandy Loam';
		}
	}
	else if((clay >= 7 && clay < 27) && (silt >= 28 && silt < 50) && (sand <= 52))
	{
		document.info.texture.value = 'Loam';
	}
	else if((silt >= 50 && (clay >= 12 && clay < 27)) || ((silt >= 50 && silt < 80) && clay < 12))
	{
		document.info.texture.value = 'Silt Loam';
	}
	else if(silt >= 80 && clay < 12)
	{
		document.info.texture.value = 'Silt';
	}
	else if((clay >= 20 && clay < 35) && (silt < 28) && (sand > 45)) 	
	{
		document.info.texture.value = 'Sandy Clay Loam';
	}
	else if((clay >= 27 && clay < 40) && (sand > 20 && sand <= 45))
	{
		document.info.texture.value = 'Clay Loam';
	}
	else if((clay >= 27 && clay < 40) && (sand  <= 20))
	{
		document.info.texture.value = 'Silty Clay Loam';
	}
	else if(clay >= 35 && sand > 45)
	{
		document.info.texture.value = 'Sandy Clay';
	}
	else if(clay >= 40 && silt >= 40)
	{
		document.info.texture.value = 'Silty Clay';
	}
	else if(clay >= 40 && sand <= 45 && silt < 40)
	{
		document.info.texture.value = 'Clay';
	}
	else
	{
		document.info.texture.value = 'Undefined';
	}
}

function getSand(vcs, cs, ms, fs, vfs)
{
	vcs = parseFloat(vcs);
	cs = parseFloat(cs);
	ms = parseFloat(ms);
	fs = parseFloat(fs);
	vfs = parseFloat(vfs);
	
	vcscs = parseFloat(vcs) + parseFloat(cs);
	vcscsms = parseFloat(vcs) + parseFloat(cs) + parseFloat(ms);
	fsvfs = parseFloat(fs) + parseFloat(vfs);

	if(vcscs >= 25 && ms < 50 && fs < 50 && vfs < 50)
	{
		document.info.texture.value = 'Coarse Sand'; 
	}
	else if((vcscsms >= 25 && vcscs < 25 && fs < 50 && vfs < 50) || (ms >= 50 && vcscs >= 25))
	{
		document.info.texture.value = 'Sand';
	}
	else if((fs >= 50 && fs > vfs) || (vcscsms < 25 && vfs < 50))
	{
		document.info.texture.value = 'Fine Sand';
	}
	else if(vfs >= 50)
	{
		document.info.texture.value = 'Very Fine Sand';
	}
	else
	{
		document.info.texture.value = 'Error';
	}
}

function getLoamSand(vcs, cs, ms, fs, vfs)
{
	vcs = parseFloat(vcs);
	cs = parseFloat(cs);
	ms = parseFloat(ms);
	fs = parseFloat(fs);
	vfs = parseFloat(vfs);

	vcscs = parseFloat(vcs) + parseFloat(cs);
	vcscsms = parseFloat(vcs) + parseFloat(cs) + parseFloat(ms);
	fsvfs = parseFloat(fs) + parseFloat(vfs);

	if(vcscs >= 25 && ms < 50 && fs < 50 && vfs < 50)
	{
		document.info.texture.value = 'Loamy Coarse Sand';
	}
	else if(vcscsms >= 25 && vcscs < 25 && fs < 50 && vfs < 50)
	{
		document.info.texture.value = 'Loamy Sand';
	}
	else if(ms >= 50 && vcscs >= 25)
	{
		document.info.texture.value = 'Loamy Sand';
	}
	else if(fs >= 50)
	{
		document.info.texture.value = 'Loamy Fine Sand';
	}
	else if(vcscsms < 25 && vfs < 50)
	{
		document.info.texture.value = 'Loamy Fine Sand';
	}
	else if(vfs >= 50)
	{
		document.info.texture.value = 'Loamy Very Fine Sand';
	}
	else
	{
		document.info.texture.value = 'Error';
	}
		
}

function getSandLoam(vcs, cs, ms, fs, vfs)
{
	vcs = parseFloat(vcs);
	cs = parseFloat(cs);
	ms = parseFloat(ms);
	fs = parseFloat(fs);
	vfs = parseFloat(vfs);

	vcscs = parseFloat(vcs) + parseFloat(cs);
	vcscsms = parseFloat(vcs) + parseFloat(cs) + parseFloat(ms);
	fsvfs = parseFloat(fs) + parseFloat(vfs);

	if(vcscs >= 25 && ms < 50 && fs < 50 && vfs < 50)
	{
		document.info.texture.value = 'Coarse Sandy Loam';
	}
	else if(vcscsms >= 30 && vfs >= 30 && vfs < 50)
	{
		document.info.texture.value = 'Coarse Sandy Loam';
	}
	else if(vcscsms >= 30 && vcscs < 25 && fs < 30 && vfs < 30)
	{
		document.info.texture.value = 'Sandy Loam';
	}
	else if(vcscsms <= 15 && fs < 30 && vfs < 30 && fsvfs < 40)
	{
		document.info.texture.value = 'Sandy Loam';
	}
	else if(vcscs >= 25 && ms >= 50)
	{
		document.info.texture.value = 'Sandy Loam';
	}
	else if(fs >= 30 && vfs < 30 && vcscs < 25)
	{
		document.info.texture.value = 'Fine Sandy Loam';
	}
	else if(vcscsms >= 15 && vcscsms < 30 && vcscs < 25)
	{
		document.info.texture.value = 'Fine Sandy Loam';
	}		
	else if(fsvfs >= 40 && fs >= vfs && vcscsms <= 15)
	{
		document.info.texture.value = 'Fine Sandy Loam';
	}
	else if(vcscs >= 25 && fs >= 50)
	{
		document.info.texture.value = 'Fine Sandy Loam';
	}
	else if(vfs >= 30 && vcscsms < 15 && vfs > fs)
	{
		document.info.texture.value = 'Very Fine Sandy Loam';
	} 
	else if(fsvfs >= 40 && vfs > fs && vcscsms < 15)
	{
		document.info.texture.value = 'Very Fine Sandy Loam';
	}
	else if(vcscs >= 25 && vfs >= 50)
	{
		document.info.texture.value = 'Very Fine Sandy Loam';
	}
	else if(vcscsms >= 30 && vfs >= 50)
	{
		document.info.texture.value = 'Very Fine Sandy Loam';
	}
	else
	{
		document.info.texture.value = 'Error';
	}
}
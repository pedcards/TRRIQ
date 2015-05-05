stRegX(h,BS="",BO=0,BT=0, ES="",ET=0, ByRef N="") {
/*	h = Haystack
	BS = beginning string
	BO = beginning offset
	BT = beginning trim, TRUE or FALSE
	ES = ending string
	ET = ending trim, TRUE or FALSE
	N = next offset
*/
	rem:="O)[OPimsxADJUXPSC(\`n)(\`r)(\`a)]+\)"
	pos0 := RegExMatch(h,((BS~=rem)?"O"BS:"O)"BS),bPat,((BO<1)?1:BO))
	pos1 := RegExMatch(h,((ES~=rem)?"O"ES:"O)"ES),ePat,pos0+bPat.len)
	N := pos1+((ET)?0:(ePat.len))
	return substr(h,pos0+((BT)?(bPat.len):0),N-pos0)
}

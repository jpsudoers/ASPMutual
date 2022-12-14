select	
	distinct H.ID_AUTORIZACION,
	h.estado,
	E.RUT as rut,
	dbo.MayMinTexto (E.R_SOCIAL) as empresa,
	CONVERT(VARCHAR(10), P.FECHA_INICIO_,	105) as FECHA_INICIO,
	dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre,
	H.ID_BLOQUE,
	CONVERT(date,	P.FECHA_INICIO_),	
	A.DOCUMENTO_COMPROMISO as doc
from
	HISTORICO_CURSOS H
inner join AUTORIZACION A on
	A.ID_AUTORIZACION = H.ID_AUTORIZACION
inner join EMPRESAS E on
	E.ID_EMPRESA = H.ID_EMPRESA
inner join PROGRAMA P on
	P.ID_PROGRAMA = H.ID_PROGRAMA
inner join CURRICULO C on
	C.ID_MUTUAL = P.ID_MUTUAL
where
	H.ESTADO = 2
	AND (CASE
		WHEN A.CON_OTIC = 1 AND A.CON_FRANQUICIA = 1 AND A.TIPO_DOC = 4 AND (A.N_REG_SENCE IS NULL OR A.N_REG_SENCE IS NOT NULL) then '0'
		WHEN A.CON_OTIC = 0 AND A.CON_FRANQUICIA = 1 AND A.TIPO_DOC = 4 AND A.N_REG_SENCE IS NOT NULL then '0'
		WHEN A.CON_OTIC = 0 AND A.CON_FRANQUICIA = 0 AND A.TIPO_DOC = 4 then '0'
		WHEN A.CON_OTIC = 0 AND A.CON_FRANQUICIA = 1 AND A.TIPO_DOC <> 4 then '1'
		WHEN A.CON_OTIC = 0	AND A.CON_FRANQUICIA = 0 AND A.TIPO_DOC <> 4 then '1'
		WHEN A.CON_OTIC = 1 AND A.CON_FRANQUICIA = 1 AND A.TIPO_DOC <> 4 AND A.N_REG_SENCE IS NOT NULL then '1'
		END)= 1
	and H.ID_PROGRAMA in (
		select
			p2.ID_PROGRAMA
		from
			PROGRAMA p2
		where CONVERT(date, p2.FECHA_INICIO_)>CONVERT(date,'25-04-2011')
		)
	--and ((isnull(e.CON_REFERENCIA, 0)= 1 and isnull(a.ID_TIPO_REFERENCIA,0)>0 and isnull(a.N_REFERENCIA, '')<> '' and a.TIPO_DOC = 0) 
	--		OR (isnull(e.CON_REFERENCIA, 0)= 1 and a.TIPO_DOC <> 0) or isnull(e.CON_REFERENCIA, 0)= 0)
union all
select	
	a.ID_AUTORIZACION,
	a.estado,
	E.RUT as rut,
	dbo.MayMinTexto (E.R_SOCIAL) as empresa,
	CONVERT(VARCHAR(10), P.FECHA_INICIO_,105) as FECHA_INICIO,
	dbo.MayMinTexto (C.NOMBRE_CURSO) as nombre,
	a.ID_BLOQUE,
	CONVERT(date,	P.FECHA_INICIO_),
	A.DOCUMENTO_COMPROMISO
from
	autorizacion a
inner join EMPRESAS E on
	E.ID_EMPRESA = a.ID_EMPRESA
inner join PROGRAMA P on
	P.ID_PROGRAMA = a.ID_PROGRAMA
inner join CURRICULO C on
	C.ID_MUTUAL = P.ID_MUTUAL
where a.ESTADO = 1
and (a.SOLO_CERTIFICADOS = 1 or a.SOLO_CERTIFICADOS is null)
and (
	select
		count(*)
	from
		HISTORICO_CURSOS hc
	where
		hc.ID_AUTORIZACION = a.ID_AUTORIZACION
		and hc.ESTADO = 0
	)= a.N_PARTICIPANTES
AND (CASE
		WHEN A.CON_OTIC = 1 AND A.CON_FRANQUICIA = 1 AND A.TIPO_DOC = 4 AND (A.N_REG_SENCE IS NULL OR A.N_REG_SENCE IS NOT NULL) then '0'
		WHEN A.CON_OTIC = 0 AND A.CON_FRANQUICIA = 1 AND A.TIPO_DOC = 4 AND A.N_REG_SENCE IS NOT NULL then '0'
		WHEN A.CON_OTIC = 0 AND A.CON_FRANQUICIA = 0 AND A.TIPO_DOC = 4 then '0'
		WHEN A.CON_OTIC = 0 AND A.CON_FRANQUICIA = 1 AND A.TIPO_DOC <> 4 then '1'
		WHEN A.CON_OTIC = 0 AND A.CON_FRANQUICIA = 0 AND A.TIPO_DOC <> 4 then '1'
		WHEN A.CON_OTIC = 1 AND A.CON_FRANQUICIA = 1 AND A.TIPO_DOC <> 4 AND A.N_REG_SENCE IS NOT NULL then '1'
		END)= 1
and a.ID_PROGRAMA in (
	select
		p2.ID_PROGRAMA
	from
		PROGRAMA p2
	where CONVERT(date,	p2.FECHA_INICIO_)>CONVERT(date,'25-04-2011')
	)
--and ((isnull(e.CON_REFERENCIA,	0)= 1 and isnull(a.ID_TIPO_REFERENCIA,0)>0 and isnull(a.N_REFERENCIA,'')<> '' and a.TIPO_DOC = 0)
--	OR (isnull(e.CON_REFERENCIA,0)= 1 and a.TIPO_DOC <> 0) or isnull(e.CON_REFERENCIA, 0)= 0)
ORDER BY CONVERT(date,	P.FECHA_INICIO_)
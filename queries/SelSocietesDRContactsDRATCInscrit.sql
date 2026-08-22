SELECT DISTINCT
  TBLSOCIETES.SocNum,
  TBLSOCIETES.SocRaison,
  TBLSOCIETES.SocTypCode,
  TBLSOCIETES.SocAdresse,
  TBLVILLES.VilCodePostal AS SocVilCodePostal,
  TBLVILLES.VilNom AS SocVilNom,
  First(SelContactsInscrit.ConNum) AS SDDRConNum,
  First(SelContactsInscrit.ConCivCode) AS SDDRConCivCode,
  First(SelContactsInscrit.ConFonCode) AS SDDRConFonCode,
  First(SelContactsInscrit.ConNom) AS SDDRConNom,
  First(SelContactsInscrit.ConPrenom) AS SDDRConPrenom,
  First(SelContactsInscrit.ConAdresse) AS SDDRConAdresse,
  First(TBLVILLES_1.VilCodePostal) AS SDDRVilCodePostal,
  First(TBLVILLES_1.VilNom) AS SDDRVilNom
FROM
  TBLVILLES
  INNER JOIN (
    (
      (
        (
          TBLSOCIETES
          INNER JOIN SelContactsInscrit ON TBLSOCIETES.SocNum = SelContactsInscrit.ConSocNum
        )
        INNER JOIN SelContactsInscrit AS SelContactsInscrit_1 ON TBLSOCIETES.SocNum = SelContactsInscrit_1.ConSocNum
      )
      LEFT JOIN TBLVILLES AS TBLVILLES_1 ON SelContactsInscrit.ConVilNum = TBLVILLES_1.VilNum
    )
    INNER JOIN TBLCONTACTS ON TBLSOCIETES.SocNum = TBLCONTACTS.ConSocNum
  ) ON TBLVILLES.VilNum = TBLSOCIETES.SocVilNum
WHERE
  (
    (
      (SelContactsInscrit.ConFonCode) = "SDDR"
    )
    AND (
      (TBLSOCIETES.SocTypCode) = "DR"
    )
    AND (
      (
        SelContactsInscrit_1.ConFonCode
      ) = "ATC"
    )
  )
GROUP BY
  TBLSOCIETES.SocNum,
  TBLSOCIETES.SocRaison,
  TBLSOCIETES.SocTypCode,
  TBLSOCIETES.SocAdresse,
  TBLVILLES.VilCodePostal,
  TBLVILLES.VilNom;

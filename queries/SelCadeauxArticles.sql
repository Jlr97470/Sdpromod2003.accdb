SELECT
  CadAchSocNum,
  CadAchNum,
  CadArtRef,
  CadArtDescription,
  CadAchDate,
  CadAchDateSaisie,
  CadAchEtat,
  CadAchQuantite,
  CadAchPoints
FROM
  TBLCADEAUXARTICLES
  INNER JOIN TBLCADEAUXACHATS ON TBLCADEAUXARTICLES.CadArtNum = TBLCADEAUXACHATS.CadAchCadArtNum
ORDER BY
  CadAchDateSaisie DESC,
  CadAchNum DESC;

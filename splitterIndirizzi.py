#Valida su Indirizzi presi dal sito del ministero

#([A-Z][\.]?[A-Z]*)[\s]+([0-9]?[\s]?[A-Z\.\'\s]*)\,?\s?([0-9]{0,3})\,?\s?([0-9]{5})\s([A-Z\'\s]*)\s\(([A-Z]{2})

#Valida su Indirizzi presi dai file di MV

#([A-Z][\.]?[A-Z]*)[\s]+([0-9]?[\s]?[A-Z\.\'\s]*)[\,]?\s+([0-9]{0,3})
#([A-Z][\.]?[A-Z]*)[\s]+([0-9]?[\s]?[A-Z\.\'\s]*)[\,]?\s+(S\.N\.C|SNC|SN|S\.N\.)



########################################################################################################################

function addressSplitting(x,y) {
  var string = x;
  const reg2 = /([A-Z][\.]?[A-Z]*)[\s]+([0-9]?[\s]?[A-Z\.\'\s]*)[\,]?\s*(S\.N\.C|SNC|SN|S\.N\.)/;
  const reg1 = /([A-Z][\.]?[A-Z]*)[\s]+([0-9]?[\s]?[A-Z\.\'\s]*)[\,]?\s*([0-9]{0,3})/;

  var found = "";

  if (x.includes("SNC") || x.includes("S.N.C")) {
    found = string.match(reg2);
  } else {
      found = string.match(reg1);
  }
  return found[y];
}

#addressSplitting("VIA BRODOLINI 14",1)

const { google } = require('googleapis');

const addHeadings = (people, headings) => {
    return people.map(personAsArray => {
      const personAsObj = {};
  
      headings.forEach((heading, i) => {
        personAsObj[heading] = personAsArray[i];
      });
  
      return personAsObj;
    });
}

const normalizeString = (value) => {
    return String(value || '')
      .trim()
      .toLowerCase();
}

const sheetValuesToObject = (sheetValues, headers) => {
    const headings = headers || sheetValues[0].map(normalizeString);
    let people = null;
    if (sheetValues) people = headers ? sheetValues : sheetValues.slice(1);
    const peopleWithHeadings = addHeadings(people, headings);
    return peopleWithHeadings;
}

let jwtClient = new google.auth.JWT(
    'siacservice-748@siac-414720.iam.gserviceaccount.com',
    null,
    '-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDM84JCn9ICthoP\n6wQgWXjk1XYPlJesccdbzy9lCAW+ofEU2HcWmjRofsxurRl2nC1wbjX+SZ3aPWqk\nFE8e9XVbuuU/5mxnfbSfrfsu7K/am21cOAqnxfqFvpfJxfmN9eP2asAlG8Xhv8dY\nNAbDhpP3yUGyr/VlWhgm6973GbpZatoDA4JuurpRxntoTJgbT3IZsXV8oNq452M+\nrJYE/H8CngVEVMly2jXPcfFKW/zm63sKkeLgRS9ljdDJEY9Cm5QaVjXqRutedShf\nsoXrPyo6/cbYG79In0l4KoCRSccaTPmSWpp/fYxokDHU0yykFS6BQF0S26huUb9m\nxC2oa+jRAgMBAAECggEAIxR9E81KQ4eK33WhD65w2G5HFnSfH3ecIXVNjfP5E7+h\nRamlfJtpZAdsE3eSd7BUwL1EhaFxIOVJRwto3YozC7/qNG1K2c306C394/LESN04\ns0OuYzgqYNEWJwW5aNiRK+DqwY9G6BSM2OaSz83NruhmU+D1VmM85hCSaeYf9OTb\n1BzDALl2SGXdsQANYrb2Pbkhvxx1THVUpHVax5MRpWMXmDs7naHy0nWKslgZ1KSZ\nT0e954Pk9Iu3XxrHeHqJcGEGbZvC0YvpiONHSzc28poNrKB/KtfgrCAoFecSc+DJ\nNch6a251ag1QexUD5Bg3REEIVn+UdISq6uTpbZx4twKBgQD+NmGWM+jyY51ii7QJ\nd62rVN2NSOv+15U9no00dYmtgL9W/e8c+6VmWgMWlNpwgFgis6Ssu4SMqW8jsIXU\nHyh6o6nkTqO3cKzz2JBLx3b7M3QmZFP1DUYR8pzYfOxS9jYyY0J7o58iw/fEI/g5\nmZtUbZmMdWYY7taahvJowrEXDwKBgQDOZHNLvOzwOY8v3mqEd1PXkpm/RwsDFuFF\nBXBKC19ka+7WuuefajJS6Z7MeJPKTGO3vHJhzn3q8YG8FQ/mYRoHcqs0dj19OXPs\ncHnqHcqH1P/Y39A9FahlSkRcTKl8m9+d1HFnrEEqLRjPW5U4DPMd8tPW75h6fRIQ\nYRKhwHQCHwKBgDznIGgQ1a1Eik8ysxZVksjqUw3nO4rZcUrK8n9v7WUg5DZeLewe\nqdiklfrR/KdZSERAD6LGZhIhAZxmTRmtwU/oZ+pnoLdxCi59YsyU4/94q0oLXUXn\nQTNJkaQYAbI6hG978lCWuahllLVr/KsoDtuiSlgpRCWTCt0ImYjZo/2nAoGBAMEA\ndES9/f/Cg6Cq54bKI5AyWi3hnG2eJrgpltDXA7RfrjAFBfYwE7EvID1rACEsAA/g\nXEIUG/HpN32PYJf586JFW84qR+PjJwFSSN9iTnNo/ntrCEsnBpr5sSVy1wdcp+bq\ns8XT8fgjxdCafta0XWCDJBAZa8gXTx4b+JVj59fXAoGAFkneRw2Yg0JjvN6bP+l1\neAiIL2862JbC+n7jeh4P0CygBH+TpoOZj15AoZx0VW97S1g/OWFzQreDZXVDqLl3\nFT7nuXCAKJrPZ50R18cUGV70Z9QaOAj79+viXwM+rPhQxY0jA4DILTZmNkES6HhJ\nHk+NAbqy1DB04u4OlU28Xzg=\n-----END PRIVATE KEY-----\n',
    ['https://www.googleapis.com/auth/spreadsheets']);
  
    //authenticate request
    jwtClient.authorize(function (err, tokens) {
          if (err) {
          console.log(err);
          return;
    } else {
         console.log("Successfully connected!");
    }
});

module.exports = {
    addHeadings,
    normalizeString,
    sheetValuesToObject,
    jwtClient
}


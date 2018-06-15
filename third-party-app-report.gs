var DOMAIN = 'FILL-ME-IN';

var HIGH_RISK_ACCESS = [
    "https://mail.google.com",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.insert",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.settings.basic",
    "https://www.googleapis.com/auth/gmail.settings.sharing",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive.metadata",
    "https://www.googleapis.com/auth/drive.photos.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/drive.scripts",
    "https://www.googleapis.com/auth/ediscovery",
    "https://www.googleapis.com/auth/ediscovery.readonly",
    "https://www.googleapis.com/auth/admin.directory.customer",
    "https://www.googleapis.com/auth/admin.directory.customer.readonly",
    "https://www.googleapis.com/auth/admin.directory.device.chromeos",
    "https://www.googleapis.com/auth/admin.directory.device.chromeos.readonly",
    "https://www.googleapis.com/auth/admin.directory.device.mobile",
    "https://www.googleapis.com/auth/admin.directory.device.mobile.action",
    "https://www.googleapis.com/auth/admin.directory.device.mobile.readonly",
    "https://www.googleapis.com/auth/admin.directory.domain",
    "https://www.googleapis.com/auth/admin.directory.domain.readonly",
    "https://www.googleapis.com/auth/admin.directory.group",
    "https://www.googleapis.com/auth/admin.directory.group.member",
    "https://www.googleapis.com/auth/admin.directory.group.member.readonly",
    "https://www.googleapis.com/auth/admin.directory.group.readonly",
    "https://www.googleapis.com/auth/admin.directory.notifications",
    "https://www.googleapis.com/auth/admin.directory.orgunit",
    "https://www.googleapis.com/auth/admin.directory.orgunit.readonly",
    "https://www.googleapis.com/auth/admin.directory.resource.calendar",
    "https://www.googleapis.com/auth/admin.directory.resource.calendar.readonly",
    "https://www.googleapis.com/auth/admin.directory.rolemanagement",
    "https://www.googleapis.com/auth/admin.directory.rolemanagement.readonly",
    "https://www.googleapis.com/auth/admin.directory.user",
    "https://www.googleapis.com/auth/admin.directory.user.alias",
    "https://www.googleapis.com/auth/admin.directory.user.alias.readonly",
    "https://www.googleapis.com/auth/admin.directory.user.readonly",
    "https://www.googleapis.com/auth/admin.directory.user.security",
    "https://www.googleapis.com/auth/admin.directory.userschema",
    "https://www.googleapis.com/auth/admin.directory.userschema.readonly",
    "https://www.googleapis.com/auth/admin.reports.audit.readonly",
    "https://www.googleapis.com/auth/admin.reports.usage.readonly"
];

//Get all users within "DOMAIN"
function listAllUsers(cb) {
    var pageToken, page;
    do {
        page = AdminDirectory.Users.list({
            domain: DOMAIN,
            orderBy: 'givenName',
            maxResults: 500,
            pageToken: pageToken
        });

        var users = page.users;
        if (users) {
            for (var i = 0; i < users.length; i++) {
                var user = users[i];
                if (cb) {
                    cb(user)
                }
            }
        } else {
            Logger.log('No users found.');
        }
        pageToken = page.nextPageToken;
    } while (pageToken);
}

//Gets all users and tokens
function step1() {
    var tokens = []
    tokens.push([
        'primaryEmail',
        'displayText',
        'clientId',
        'anonymous',
        'scopes'
    ]);

    listAllUsers(function(user) {
        try {
            if (user.suspended) {
                Logger.log('[suspended] %s (%s)', user.name.fullName, user.primaryEmail);
                return;
            }

            var currentTokens = AdminDirectory.Tokens.list(user.primaryEmail);
            if (currentTokens && currentTokens.items && currentTokens.items.length) {
                for (var i = 0; i < currentTokens.items.length; i++) {
                    var tok = currentTokens.items[i];
                    if (tok.nativeApp == false) {
                        tokens.push([
                            user.primaryEmail,
                            tok.displayText,
                            tok.clientId,
                            tok.anonymous,
                            tok.scopes.join('\n'),
                        ]);
                    }
                }
            }
        } catch (e) {
            Logger.log("[error] %s: %s", user.primaryEmail, e);
        }
    });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users")
    if (sheet == null) {
        sheet = ss.insertSheet("Users");
    } else {
        sheet.clear();
    }

    Logger.log('Tokens written to Sheet Users: %s', tokens.length);
    var dataRange = sheet.getRange(1, 1, tokens.length, tokens[0].length);
    dataRange.setValues(tokens);
}

//Get counts of token usage
function step2() {
    var countsRows = [];
    countsRows.push([
        "numInstalls",
        "displayText",
        "clientId",
        "highRisk",
        "scopes"
    ])

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users")
    if (sheet == null) {
        Logger.log('Did not find Users tab. First run function "step1"')
        return;
    }
    var range = sheet.getDataRange();
    var tokens = range.getValues();
    tokens.shift(); //Remove header

    //Get counts of each token. Format [clientId = count]
    Logger.log('Counting tokens...');
    var tokenInstallCount = tokens.reduce(function(sums, entry) {
        sums[entry[2]] = (sums[entry[2]] || 0) + 1;
        return sums;
    }, {});


    Logger.log('Retrieving information associated with clientId...');
    for (tokenRow in tokenInstallCount) {
        //Retrieve information associated with clientId
        var token = [];
        for (var i = 0; i < tokens.length; i++) {
            if (tokens[i][2] == tokenRow) {
                token = tokens[i];
                break;
            }
        }
        if (token == null) {
            Logger.log("Error: token not found");
            return;
        }
        //Check if scopes appear in HIGH_RISK_ACCESS
        var match = false;
        oauth_scopes = token[4].split('\n');
        if (HIGH_RISK_ACCESS.some(function(element) {
                //Logger.log(token);
                return oauth_scopes.indexOf(element) >= 0;
            })) {
            match = true;
        }

        countsRows.push([
            tokenInstallCount[tokenRow],
            token[1],
            token[2],
            match,
            token[4]
        ])
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Counts")
    if (sheet == null) {
        sheet = ss.insertSheet("Counts");
    } else {
        sheet.clear();
    }

    var dataRange = sheet.getRange(1, 1, countsRows.length, countsRows[0].length);
    dataRange.setValues(countsRows);
    sheet.sort(1, false);

    Logger.log('Finished');

}

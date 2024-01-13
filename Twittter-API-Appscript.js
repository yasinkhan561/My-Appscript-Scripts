function Twitter_get_user_info(
	string_Screen_name,
	string_Consumer_key,
	string_Consumer_secret
) {
	var spreadsheet_Twitter_user_data =
		SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Twitter user data");
	spreadsheet_Twitter_user_data.getRange(3, 1, 1, 20).clearContent();

	var tokenUrl = "https://api.twitter.com/oauth2/token";
	var tokenCredential = Utilities.base64EncodeWebSafe(
		string_Consumer_key + ":" + string_Consumer_secret
	);

	var tokenOptions = {
		headers: {
			Authorization: "Basic " + tokenCredential,
			"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
		},
		method: "post",
		payload: "grant_type=client_credentials",
	};

	var responseToken = UrlFetchApp.fetch(tokenUrl, tokenOptions);
	var parsedToken = JSON.parse(responseToken);
	var token = parsedToken.access_token;
	var apiUrl = "";
	var responseApi = "";

	var apiOptions = {
		headers: {
			Authorization: "Bearer " + token,
		},
		method: "get",
	};

	var string_Column_a = "";
	var string_User_profile_image = "";
	var string_User_screen_name = "";
	var string_User_name = "";
	var string_Location = "";
	var string_Created_at = "";
	var string_Followers_count = "";
	var string_User = "";
	var string_Favourites_count = "";
	var string_Language = "";
	var string_Protected = "";
	var string_Time_zone = "";
	var string_Verified = "";
	var string_Statuses_count = "";
	var string_Url = "";
	var string_Description = "";

	var string_Next_cursor = -1;

	apiUrl =
		"https://api.twitter.com/1.1/users/show.json?screen_name=" +
		string_Screen_name;
	responseApi = UrlFetchApp.fetch(apiUrl, apiOptions);

	if (responseApi.getResponseCode() == 200) {
		var obj_data = JSON.parse(responseApi.getContentText());

		string_Column_a = "1";
		string_User_profile_image = '=IMAGE("' + obj_data.profile_image_url + '")';
		string_User_screen_name = obj_data.screen_name;
		string_User_name = obj_data.name;
		string_Location = obj_data.location;
		string_Created_at = obj_data.created_at;
		string_Followers_count = obj_data.followers_count;
		string_User = obj_data.friends_count;
		string_Favourites_count = obj_data.favourites_count;
		string_Language = obj_data.lang;
		string_Protected = obj_data.protected;
		string_Time_zone = obj_data.time_zone;
		string_Verified = obj_data.verified;
		string_Statuses_count = obj_data.statuses_count;
		string_Url = obj_data.url;
		string_Description = obj_data.description;
	}

	spreadsheet_Twitter_user_data.getRange("A3").setValue(string_Column_a);
	spreadsheet_Twitter_user_data
		.getRange("B3")
		.setValue(string_User_profile_image);
	spreadsheet_Twitter_user_data
		.getRange("C3")
		.setValue(string_User_screen_name);
	spreadsheet_Twitter_user_data.getRange("D3").setValue(string_User_name);
	spreadsheet_Twitter_user_data.getRange("E3").setValue(string_Location);
	spreadsheet_Twitter_user_data.getRange("F3").setValue(string_Created_at);
	spreadsheet_Twitter_user_data.getRange("G3").setValue(string_Followers_count);
	spreadsheet_Twitter_user_data.getRange("H3").setValue(string_User);
	spreadsheet_Twitter_user_data
		.getRange("I3")
		.setValue(string_Favourites_count);
	spreadsheet_Twitter_user_data.getRange("J3").setValue(string_Language);
	spreadsheet_Twitter_user_data.getRange("K3").setValue(string_Protected);
	spreadsheet_Twitter_user_data.getRange("L3").setValue(string_Time_zone);
	spreadsheet_Twitter_user_data.getRange("M3").setValue(string_Verified);
	spreadsheet_Twitter_user_data.getRange("N3").setValue(string_Statuses_count);
	spreadsheet_Twitter_user_data.getRange("O3").setValue(string_Url);
	spreadsheet_Twitter_user_data.getRange("P3").setValue(string_Description);
}

function Twitter_get_friends(
	string_Screen_name,
	string_Consumer_key,
	string_Consumer_secret
) {
	var spreadsheet_Friends =
		SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Friends");
	spreadsheet_Friends.getRange(3, 1, 2600, 20).clearContent();

	var tokenUrl = "https://api.twitter.com/oauth2/token";
	var tokenCredential = Utilities.base64EncodeWebSafe(
		string_Consumer_key + ":" + string_Consumer_secret
	);

	var tokenOptions = {
		headers: {
			Authorization: "Basic " + tokenCredential,
			"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
		},
		method: "post",
		payload: "grant_type=client_credentials",
	};

	var responseToken = UrlFetchApp.fetch(tokenUrl, tokenOptions);
	var parsedToken = JSON.parse(responseToken);
	var token = parsedToken.access_token;
	var apiUrl = "";
	var responseApi = "";

	var apiOptions = {
		headers: {
			Authorization: "Bearer " + token,
		},
		method: "get",
	};

	var array_Column_a = [];
	var array_Friends_profile_image = [];
	var array_Friends_screen_name = [];
	var array_Friends_name = [];
	var array_Location = [];
	var array_Created_at = [];
	var array_Followers_count = [];
	var array_Friends = [];
	var array_Favourites_count = [];
	var array_Language = [];
	var array_Protected = [];
	var array_Time_zone = [];
	var array_Verified = [];
	var array_Statuses_count = [];
	var array_Url = [];
	var array_Description = [];

	var string_Next_cursor = -1;
	var int_Line_counter = 1;

	do {
		apiUrl =
			"https://api.twitter.com/1.1/friends/list.json?cursor=" +
			string_Next_cursor +
			"&screen_name=" +
			string_Screen_name +
			"&skip_status=true&include_user_entities=false&count=200";
		responseApi = UrlFetchApp.fetch(apiUrl, apiOptions);

		if (responseApi.getResponseCode() == 200) {
			var obj_data = JSON.parse(responseApi.getContentText());

			for (var int_i = 0; int_i < obj_data.users.length; int_i++) {
				array_Column_a.push([int_Line_counter]);
				array_Friends_profile_image.push([
					'=IMAGE("' + obj_data.users[int_i].profile_image_url + '")',
				]);
				array_Friends_screen_name.push([obj_data.users[int_i].screen_name]);
				array_Friends_name.push([obj_data.users[int_i].name]);
				array_Location.push([obj_data.users[int_i].location]);
				array_Created_at.push([obj_data.users[int_i].created_at]);
				array_Followers_count.push([obj_data.users[int_i].followers_count]);
				array_Friends.push([obj_data.users[int_i].friends_count]);
				array_Favourites_count.push([obj_data.users[int_i].favourites_count]);
				array_Language.push([obj_data.users[int_i].lang]);
				array_Protected.push([obj_data.users[int_i].protected]);
				array_Time_zone.push([obj_data.users[int_i].time_zone]);
				array_Verified.push([obj_data.users[int_i].verified]);
				array_Statuses_count.push([obj_data.users[int_i].statuses_count]);
				array_Url.push([obj_data.users[int_i].url]);
				array_Description.push([obj_data.users[int_i].description]);

				int_Line_counter++;
			}

			string_Next_cursor = obj_data.next_cursor;
		}
	} while (string_Next_cursor != 0 && int_Line_counter < 2000);

	if (array_Column_a.length > 0) {
		spreadsheet_Friends
			.getRange("A3:A" + (array_Column_a.length + 2))
			.setValues(array_Column_a);
		spreadsheet_Friends
			.getRange("B3:B" + (array_Friends_profile_image.length + 2))
			.setValues(array_Friends_profile_image);
		spreadsheet_Friends
			.getRange("C3:C" + (array_Friends_screen_name.length + 2))
			.setValues(array_Friends_screen_name);
		spreadsheet_Friends
			.getRange("D3:D" + (array_Friends_name.length + 2))
			.setValues(array_Friends_name);
		spreadsheet_Friends
			.getRange("E3:E" + (array_Location.length + 2))
			.setValues(array_Location);
		spreadsheet_Friends
			.getRange("F3:F" + (array_Created_at.length + 2))
			.setValues(array_Created_at);
		spreadsheet_Friends
			.getRange("G3:G" + (array_Followers_count.length + 2))
			.setValues(array_Followers_count);
		spreadsheet_Friends
			.getRange("H3:H" + (array_Friends.length + 2))
			.setValues(array_Friends);
		spreadsheet_Friends
			.getRange("I3:I" + (array_Favourites_count.length + 2))
			.setValues(array_Favourites_count);
		spreadsheet_Friends
			.getRange("J3:J" + (array_Language.length + 2))
			.setValues(array_Language);
		spreadsheet_Friends
			.getRange("K3:K" + (array_Protected.length + 2))
			.setValues(array_Protected);
		spreadsheet_Friends
			.getRange("L3:L" + (array_Time_zone.length + 2))
			.setValues(array_Time_zone);
		spreadsheet_Friends
			.getRange("M3:M" + (array_Verified.length + 2))
			.setValues(array_Verified);
		spreadsheet_Friends
			.getRange("N3:N" + (array_Statuses_count.length + 2))
			.setValues(array_Statuses_count);
		spreadsheet_Friends
			.getRange("O3:O" + (array_Url.length + 2))
			.setValues(array_Url);
		spreadsheet_Friends
			.getRange("P3:P" + (array_Description.length + 2))
			.setValues(array_Description);
	} else {
		Browser.msgBox("0 Friends found.");
	}
}

function Twitter_get_followers(
	string_Screen_name,
	string_Consumer_key,
	string_Consumer_secret
) {
	var spreadsheet_Followers =
		SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Followers");
	spreadsheet_Followers.getRange(3, 1, 2600, 20).clearContent();

	var tokenUrl = "https://api.twitter.com/oauth2/token";
	var tokenCredential = Utilities.base64EncodeWebSafe(
		string_Consumer_key + ":" + string_Consumer_secret
	);

	var tokenOptions = {
		headers: {
			Authorization: "Basic " + tokenCredential,
			"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
		},
		method: "post",
		payload: "grant_type=client_credentials",
	};

	var responseToken = UrlFetchApp.fetch(tokenUrl, tokenOptions);
	var parsedToken = JSON.parse(responseToken);
	var token = parsedToken.access_token;
	var apiUrl = "";
	var responseApi = "";

	var apiOptions = {
		headers: {
			Authorization: "Bearer " + token,
		},
		method: "get",
	};

	var array_Column_a = [];
	var array_Friends_profile_image = [];
	var array_Friends_screen_name = [];
	var array_Friends_name = [];
	var array_Location = [];
	var array_Created_at = [];
	var array_Followers_count = [];
	var array_Friends = [];
	var array_Favourites_count = [];
	var array_Language = [];
	var array_Protected = [];
	var array_Time_zone = [];
	var array_Verified = [];
	var array_Statuses_count = [];
	var array_Url = [];
	var array_Description = [];

	var string_Next_cursor = -1;
	var int_Line_counter = 1;

	do {
		apiUrl =
			"https://api.twitter.com/1.1/followers/list.json?cursor=" +
			string_Next_cursor +
			"&screen_name=" +
			string_Screen_name +
			"&skip_status=true&include_user_entities=false&count=200";
		responseApi = UrlFetchApp.fetch(apiUrl, apiOptions);

		if (responseApi.getResponseCode() == 200) {
			var obj_data = JSON.parse(responseApi.getContentText());

			for (var int_i = 0; int_i < obj_data.users.length; int_i++) {
				array_Column_a.push([int_Line_counter]);
				array_Friends_profile_image.push([
					'=IMAGE("' + obj_data.users[int_i].profile_image_url + '")',
				]);
				array_Friends_screen_name.push([obj_data.users[int_i].screen_name]);
				array_Friends_name.push([obj_data.users[int_i].name]);
				array_Location.push([obj_data.users[int_i].location]);
				array_Created_at.push([obj_data.users[int_i].created_at]);
				array_Followers_count.push([obj_data.users[int_i].followers_count]);
				array_Friends.push([obj_data.users[int_i].friends_count]);
				array_Favourites_count.push([obj_data.users[int_i].favourites_count]);
				array_Language.push([obj_data.users[int_i].lang]);
				array_Protected.push([obj_data.users[int_i].protected]);
				array_Time_zone.push([obj_data.users[int_i].time_zone]);
				array_Verified.push([obj_data.users[int_i].verified]);
				array_Statuses_count.push([obj_data.users[int_i].statuses_count]);
				array_Url.push([obj_data.users[int_i].url]);
				array_Description.push([obj_data.users[int_i].description]);

				int_Line_counter++;
			}

			string_Next_cursor = obj_data.next_cursor;
		}
	} while (string_Next_cursor != 0 && int_Line_counter < 2000);

	if (array_Column_a.length > 0) {
		spreadsheet_Followers
			.getRange("A3:A" + (array_Column_a.length + 2))
			.setValues(array_Column_a);
		spreadsheet_Followers
			.getRange("B3:B" + (array_Friends_profile_image.length + 2))
			.setValues(array_Friends_profile_image);
		spreadsheet_Followers
			.getRange("C3:C" + (array_Friends_screen_name.length + 2))
			.setValues(array_Friends_screen_name);
		spreadsheet_Followers
			.getRange("D3:D" + (array_Friends_name.length + 2))
			.setValues(array_Friends_name);
		spreadsheet_Followers
			.getRange("E3:E" + (array_Location.length + 2))
			.setValues(array_Location);
		spreadsheet_Followers
			.getRange("F3:F" + (array_Created_at.length + 2))
			.setValues(array_Created_at);
		spreadsheet_Followers
			.getRange("G3:G" + (array_Followers_count.length + 2))
			.setValues(array_Followers_count);
		spreadsheet_Followers
			.getRange("H3:H" + (array_Friends.length + 2))
			.setValues(array_Friends);
		spreadsheet_Followers
			.getRange("I3:I" + (array_Favourites_count.length + 2))
			.setValues(array_Favourites_count);
		spreadsheet_Followers
			.getRange("J3:J" + (array_Language.length + 2))
			.setValues(array_Language);
		spreadsheet_Followers
			.getRange("K3:K" + (array_Protected.length + 2))
			.setValues(array_Protected);
		spreadsheet_Followers
			.getRange("L3:L" + (array_Time_zone.length + 2))
			.setValues(array_Time_zone);
		spreadsheet_Followers
			.getRange("M3:M" + (array_Verified.length + 2))
			.setValues(array_Verified);
		spreadsheet_Followers
			.getRange("N3:N" + (array_Statuses_count.length + 2))
			.setValues(array_Statuses_count);
		spreadsheet_Followers
			.getRange("O3:O" + (array_Url.length + 2))
			.setValues(array_Url);
		spreadsheet_Followers
			.getRange("P3:P" + (array_Description.length + 2))
			.setValues(array_Description);
	} else {
		Browser.msgBox("0 Followers found.");
	}
}

function Twitter_get_tweets(
	string_Screen_name,
	string_Consumer_key,
	string_Consumer_secret
) {
	var spreadsheet_Tweets =
		SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tweets");
	spreadsheet_Tweets.getRange(3, 1, 2600, 20).clearContent();

	var tokenUrl = "https://api.twitter.com/oauth2/token";
	var tokenCredential = Utilities.base64EncodeWebSafe(
		string_Consumer_key + ":" + string_Consumer_secret
	);

	var tokenOptions = {
		headers: {
			Authorization: "Basic " + tokenCredential,
			"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
		},
		method: "post",
		payload: "grant_type=client_credentials",
	};

	var responseToken = UrlFetchApp.fetch(tokenUrl, tokenOptions);
	var parsedToken = JSON.parse(responseToken);
	var token = parsedToken.access_token;
	var apiUrl = "";
	var responseApi = "";

	var apiOptions = {
		headers: {
			Authorization: "Bearer " + token,
		},
		method: "get",
	};

	var array_Column_a = [];
	var array_Created_at = [];
	var array_Text = [];
	var array_Expanded_url = [];
	var array_Media_url = [];
	var array_Friends_profile_image = [];
	var array_Screen_name = [];
	var array_Name = [];
	var array_Location = [];
	var array_User_expanded_url = [];
	var array_Statuses_count = [];

	var string_Max_id = 0;
	var int_Line_counter = 1;
	var int_Break_loop = 0;

	do {
		if (int_Line_counter == 1) {
			apiUrl =
				"https://api.twitter.com/1.1/statuses/user_timeline.json?screen_name=" +
				string_Screen_name +
				"&count=200&include_rts=1";
		} else {
			apiUrl =
				"https://api.twitter.com/1.1/statuses/user_timeline.json?screen_name=" +
				string_Screen_name +
				"&count=200&include_rts=1&max_id=" +
				string_Max_id;
		}
		responseApi = UrlFetchApp.fetch(apiUrl, apiOptions);

		if (responseApi.getResponseCode() == 200) {
			var obj_data = JSON.parse(responseApi.getContentText());

			for (var int_i = 0; int_i < obj_data.length; int_i++) {
				array_Column_a.push([int_Line_counter]);
				array_Created_at.push([obj_data[int_i].created_at]);
				array_Text.push([obj_data[int_i].text]);

				if (
					obj_data[int_i].entities.urls[0] != undefined &&
					obj_data[int_i].entities != undefined
				) {
					array_Expanded_url.push([
						obj_data[int_i].entities.urls[0].expanded_url,
					]);
				} else {
					array_Expanded_url.push([""]);
				}

				if (
					obj_data[int_i].entities.media != undefined &&
					obj_data[int_i].entities != undefined
				) {
					array_Media_url.push([
						'=IMAGE("' + obj_data[int_i].entities.media[0].media_url + '")',
					]);
				} else {
					array_Media_url.push([""]);
				}

				array_Friends_profile_image.push([
					'=IMAGE("' + obj_data[int_i].user.profile_image_url + '")',
				]);
				array_Screen_name.push([obj_data[int_i].user.screen_name]);
				array_Name.push([obj_data[int_i].user.name]);
				array_Location.push([obj_data[int_i].user.location]);

				if (
					obj_data[int_i].user.entities != undefined &&
					obj_data[int_i].user.entities.url != undefined &&
					obj_data[int_i].user.entities.url.urls != undefined
				) {
					array_User_expanded_url.push([
						obj_data[int_i].user.entities.url.urls[0].expanded_url,
					]);
				} else {
					array_User_expanded_url.push([""]);
				}
				array_Statuses_count.push([obj_data[int_i].user.statuses_count]);

				int_Line_counter++;
			}

			if (
				obj_data[obj_data.length - 1] != undefined &&
				int_i < parseInt(obj_data[0].user.statuses_count)
			) {
				string_Max_id = obj_data[obj_data.length - 1].id;
			} else {
				int_Break_loop = 1;
			}
		}
	} while (int_Break_loop != 1 && int_Line_counter < 1000);

	if (array_Column_a.length > 0) {
		spreadsheet_Tweets
			.getRange("A3:A" + (array_Column_a.length + 2))
			.setValues(array_Column_a);
		spreadsheet_Tweets
			.getRange("B3:B" + (array_Created_at.length + 2))
			.setValues(array_Created_at);
		spreadsheet_Tweets
			.getRange("C3:C" + (array_Text.length + 2))
			.setValues(array_Text);
		spreadsheet_Tweets
			.getRange("D3:D" + (array_Expanded_url.length + 2))
			.setValues(array_Expanded_url);
		spreadsheet_Tweets
			.getRange("E3:E" + (array_Media_url.length + 2))
			.setValues(array_Media_url);
		spreadsheet_Tweets
			.getRange("F3:F" + (array_Friends_profile_image.length + 2))
			.setValues(array_Friends_profile_image);
		spreadsheet_Tweets
			.getRange("G3:G" + (array_Screen_name.length + 2))
			.setValues(array_Screen_name);
		spreadsheet_Tweets
			.getRange("H3:H" + (array_Name.length + 2))
			.setValues(array_Name);
		spreadsheet_Tweets
			.getRange("I3:I" + (array_Location.length + 2))
			.setValues(array_Location);
		spreadsheet_Tweets
			.getRange("J3:J" + (array_User_expanded_url.length + 2))
			.setValues(array_User_expanded_url);
		spreadsheet_Tweets
			.getRange("K3:K" + (array_Statuses_count.length + 2))
			.setValues(array_Statuses_count);
	} else {
		Browser.msgBox("0 Tweets found");
	}
}

function Get_User_info() {
	var spreadsheet_Options =
		SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Options");

	var string_Screen_name = spreadsheet_Options.getRange("C3").getDisplayValue();
	var string_Consumer_key = spreadsheet_Options
		.getRange("C4")
		.getDisplayValue();
	var string_Consumer_secret = spreadsheet_Options
		.getRange("C5")
		.getDisplayValue();

	Twitter_get_user_info(
		string_Screen_name,
		string_Consumer_key,
		string_Consumer_secret
	);
}

function Get_friends_data() {
	var spreadsheet_Options =
		SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Options");

	var string_Screen_name = spreadsheet_Options.getRange("C3").getDisplayValue();
	var string_Consumer_key = spreadsheet_Options
		.getRange("C4")
		.getDisplayValue();
	var string_Consumer_secret = spreadsheet_Options
		.getRange("C5")
		.getDisplayValue();

	Twitter_get_friends(
		string_Screen_name,
		string_Consumer_key,
		string_Consumer_secret
	);
}

function Get_followers_data() {
	var spreadsheet_Options =
		SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Options");

	var string_Screen_name = spreadsheet_Options.getRange("C3").getDisplayValue();
	var string_Consumer_key = spreadsheet_Options
		.getRange("C4")
		.getDisplayValue();
	var string_Consumer_secret = spreadsheet_Options
		.getRange("C5")
		.getDisplayValue();

	Twitter_get_followers(
		string_Screen_name,
		string_Consumer_key,
		string_Consumer_secret
	);
}

function Get_tweets_data() {
	var spreadsheet_Options =
		SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Options");

	var string_Screen_name = spreadsheet_Options.getRange("C3").getDisplayValue();
	var string_Consumer_key = spreadsheet_Options
		.getRange("C4")
		.getDisplayValue();
	var string_Consumer_secret = spreadsheet_Options
		.getRange("C5")
		.getDisplayValue();

	Twitter_get_tweets(
		string_Screen_name,
		string_Consumer_key,
		string_Consumer_secret
	);
}

function Run_all_the_scripts() {
	Get_User_info();
	Get_friends_data();
	Get_followers_data();
	Get_tweets_data();
}

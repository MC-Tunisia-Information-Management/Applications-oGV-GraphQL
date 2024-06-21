function getOpportunityApplications_oGV() {
  var graphQLQuery = `
      query AllOpportunityApplication {
        allOpportunityApplication(
          filters: {
            person_committee: lc_id,
            created_at: { from: "2024-02-01T00:00:00Z", to: "2025-01-31T00:00:00Z" },
            programmes: 7
          },
          pagination: { per_page: 3000 }
        ) {
          data {
            id
            person {
              full_name
              id
              phone
              person_profile {
                backgrounds {
                  name
                }
              }
            }
            created_at
            host_lc_name
            home_mc {
              name
            }
            current_status
            opportunity {
              title
              id
              opportunity_duration_type {
                duration_type
              }
              programme {
                short_name_display
              }
            }
          }
        }
      }
    `;

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ query: graphQLQuery }),
  };

  var url = "https://gis-api.aiesec.org/graphql?access_token={access_token}";
  var response = UrlFetchApp.fetch(url, options);

  // Parse response
  var responseData = JSON.parse(response.getContentText());
  var applicationData = responseData.data.allOpportunityApplication.data;

  // Get Google Sheets spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("🌎 applications | oGV");

  // Write data
  var data = applicationData.map(function (application) {
    return [
      application.id,
      application.person.full_name,
      application.person.id,
      application.person.phone,
      application.person.person_profile
        ? application.person.person_profile.backgrounds
            .map(function (background) {
              return background.name;
            })
            .join(", ")
        : "", // Handle null values
      application.created_at,
      application.host_lc_name,
      application.home_mc ? application.home_mc.name : "", // Handle null values
      application.current_status,
      application.opportunity.title,
      application.opportunity.id,
      application.opportunity.opportunity_duration_type
        ? application.opportunity.opportunity_duration_type.duration_type
        : "", // Handle null values
      application.opportunity.programme
        ? application.opportunity.programme.short_name_display
        : "", // Handle null values
    ];
  });
  sheet.getRange("B7:N" + (data.length + 6)).setValues(data); // Adjusted range
}

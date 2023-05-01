using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using RestSharp.Authenticators;

namespace Controlling
{

    public class JiraIssue
    {
        public string IssueType { get; set; }
        public string Key { get; set; }
        public string Summary { get; set; }
        public double? StoryPoints { get; set; }
        public List<string> Labels { get; set; }
        public string? Assignee { get; set; }
        public string Reporter { get; set; }
        public string Priority { get; set; }
        public string Status { get; set; }
        public string Resolution { get; set; }
        public DateTime Created { get; set; }
        public DateTime Updated { get; set; }
        public DateTime? DueDate { get; set; }
        public string? Sprint { get; set; }
        public string? EpicLink { get; set; }
        public List<string> SubTasks { get; set; }
        public string? Verrechenbarkeit { get; set; }
    }
    public class JiraApiClient
    {
        private string storyPointField = "Story Points[Number]";
        private string sprintField = "customfield_10020";
        private string chargableField = "customfield_10034";

        private readonly string baseUrl;
        private readonly string username;
        private readonly string apiToken;
        private readonly RestClient client;

        public JiraApiClient(string baseUrl, string username, string apiToken)
        {
            this.baseUrl = baseUrl;
            this.username = username;
            this.apiToken = apiToken;
            this.client = new RestClient(baseUrl);
        }

        public List<JiraIssue> GetAllIssuesWithFields(string jql)
        {
            var issues = new List<JiraIssue>();
            int startAt = 0;
            int maxResults = 50;

            //GetFields();

            while (true)
            {
                var request = new RestRequest("rest/api/3/search", Method.Get);
                request.AddHeader("Accept", "application/json");
                request.AddHeader("Content-Type", "application/json");
                request.AddHeader("Authorization", $"Basic {Convert.ToBase64String(Encoding.ASCII.GetBytes($"{username}:{apiToken}"))}");
                request.AddQueryParameter("jql", jql);
                request.AddQueryParameter("startAt", startAt.ToString());
                request.AddQueryParameter("maxResults", maxResults.ToString());

                // Replace customfield_10004, customfield_10005, and customfield_11000 with the custom field IDs for
                // "Story Points", "Sprint", and "Verrechenbarkeit" respectively. You can find these custom field IDs in your Jira instance by going to the "Custom Fields" configuration page.
                request.AddQueryParameter("fields", $"issuetype,key,summary,{storyPointField},labels,assignee,reporter,priority,status,resolution,created,updated,duedate,{sprintField},epic,subtasks,{chargableField}");
                //request.AddQueryParameter("fields", $"issuetype,key,summary");

                var response = client.Execute(request);

                if (response.IsSuccessful)
                {
                    var jsonResponse = JsonDocument.Parse(response.Content);
                    var issueArray = jsonResponse.RootElement.GetProperty("issues");
                    /*
                      jsonResponse.RootElement.GetProperty("issues")[1].GetProperty("key");
                        ValueKind = String : "LGS-573"
                            ValueKind: String
                      jsonResponse.RootElement.GetProperty("issues")[1].GetProperty("fields").GetProperty("summary");
                        ValueKind = String : "Benutzerinformationen: excluded for billing checkbox"
                            ValueKind: String
                     */
                    foreach (var issue in issueArray.EnumerateArray())
                    {
                        var jiraIssue = ParseJiraIssue(issue);
                        issues.Add(jiraIssue);
                    }

                    int total = jsonResponse.RootElement.GetProperty("total").GetInt32();

                    if (startAt + maxResults >= total)
                    {
                        break;
                    }
                    else
                    {
                        startAt += maxResults;
                    }
                }
                else
                {
                    throw new Exception($"Failed to fetch issues: {response.ErrorMessage}");
                }
            }

            return issues;
        }

        private void GetFields()
        {
            //var request = new RestRequest("rest/api/3/project", Method.Get);
            var request = new RestRequest("rest/api/3/field", Method.Get);
            request.AddHeader("Accept", "application/json");
            request.AddHeader("Content-Type", "application/json");
            request.AddHeader("Authorization", $"Basic {Convert.ToBase64String(Encoding.ASCII.GetBytes($"{username}:{apiToken}"))}");
            //request.Authenticator = new HttpBasicAuthenticator { }
            var response = client.Execute(request);
            // [{"id":"statuscategorychangedate","key":"statuscategorychangedate","name":"Status Category Changed","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["statusCategoryChangedDate"],"schema":{"type":"datetime","system":"statuscategorychangedate"}},{"id":"parent","key":"parent","name":"Parent","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["parent"]},{"id":"issuekey","key":"issuekey","name":"Key","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["id","issue","issuekey","key"]},{"id":"timespent","key":"timespent","name":"Time Spent","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["timespent"],"schema":{"type":"number","system":"timespent"}},{"id":"timeoriginalestimate","key":"timeoriginalestimate","name":"Original estimate","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["originalEstimate","timeoriginalestimate"],"schema":{"type":"number","system":"timeoriginalestimate"}},{"id":"project","key":"project","name":"Project","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["project"],"schema":{"type":"project","system":"project"}},{"id":"aggregatetimespent","key":"aggregatetimespent","name":"Σ Time Spent","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":[],"schema":{"type":"number","system":"aggregatetimespent"}},{"id":"statusCategory","key":"statusCategory","name":"Status Category","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["statusCategory"]},{"id":"timetracking","key":"timetracking","name":"Time tracking","custom":false,"orderable":true,"navigable":false,"searchable":true,"clauseNames":[],"schema":{"type":"timetracking","system":"timetracking"}},{"id":"attachment","key":"attachment","name":"Attachment","custom":false,"orderable":true,"navigable":false,"searchable":true,"clauseNames":["attachments"],"schema":{"type":"array","items":"attachment","system":"attachment"}},{"id":"aggregatetimeestimate","key":"aggregatetimeestimate","name":"Σ Remaining Estimate","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":[],"schema":{"type":"number","system":"aggregatetimeestimate"}},{"id":"resolutiondate","key":"resolutiondate","name":"Resolved","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["resolutiondate","resolved"],"schema":{"type":"datetime","system":"resolutiondate"}},{"id":"workratio","key":"workratio","name":"Work Ratio","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["workratio"],"schema":{"type":"number","system":"workratio"}},{"id":"issuerestriction","key":"issuerestriction","name":"Restrict to","custom":false,"orderable":true,"navigable":false,"searchable":true,"clauseNames":[],"schema":{"type":"issuerestriction","system":"issuerestriction"}},{"id":"lastViewed","key":"lastViewed","name":"Last Viewed","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["lastViewed"],"schema":{"type":"datetime","system":"lastViewed"}},{"id":"watches","key":"watches","name":"Watchers","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["watchers"],"schema":{"type":"watches","system":"watches"}},{"id":"thumbnail","key":"thumbnail","name":"Images","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":[]},{"id":"creator","key":"creator","name":"Creator","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["creator"],"schema":{"type":"user","system":"creator"}},{"id":"subtasks","key":"subtasks","name":"Sub-tasks","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["subtasks"],"schema":{"type":"array","items":"issuelinks","system":"subtasks"}},{"id":"created","key":"created","name":"Created","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["created","createdDate"],"schema":{"type":"datetime","system":"created"}},{"id":"aggregateprogress","key":"aggregateprogress","name":"Σ Progress","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":[],"schema":{"type":"progress","system":"aggregateprogress"}},{"id":"timeestimate","key":"timeestimate","name":"Remaining Estimate","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["remainingEstimate","timeestimate"],"schema":{"type":"number","system":"timeestimate"}},{"id":"aggregatetimeoriginalestimate","key":"aggregatetimeoriginalestimate","name":"Σ Original Estimate","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":[],"schema":{"type":"number","system":"aggregatetimeoriginalestimate"}},{"id":"progress","key":"progress","name":"Progress","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["progress"],"schema":{"type":"progress","system":"progress"}},{"id":"comment","key":"comment","name":"Comment","custom":false,"orderable":true,"navigable":false,"searchable":true,"clauseNames":["comment"],"schema":{"type":"comments-page","system":"comment"}},{"id":"votes","key":"votes","name":"Votes","custom":false,"orderable":false,"navigable":true,"searchable":false,"clauseNames":["votes"],"schema":{"type":"votes","system":"votes"}},{"id":"worklog","key":"worklog","name":"Log Work","custom":false,"orderable":true,"navigable":false,"searchable":true,"clauseNames":[],"schema":{"type":"array","items":"worklog","system":"worklog"}},{"id":"updated","key":"updated","name":"Updated","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["updated","updatedDate"],"schema":{"type":"datetime","system":"updated"}},{"id":"status","key":"status","name":"Status","custom":false,"orderable":false,"navigable":true,"searchable":true,"clauseNames":["status"],"schema":{"type":"status","system":"status"}}]
            Console.WriteLine(response.Content);
            Console.WriteLine();
        }

        private JiraIssue ParseJiraIssue(JsonElement issue)
        {
            var fields = issue.GetProperty("fields");

            return new JiraIssue
            {
                IssueType = fields.GetProperty("issuetype").GetProperty("name").GetString(),
                Key = issue.GetProperty("key").GetString(),
                Summary = fields.GetProperty("summary").GetString(),
                StoryPoints = fields.GetProperty(storyPointField).TryGetDouble(out var storyPoints) ? (double?)storyPoints : null,
                Labels = fields.GetProperty("labels").EnumerateArray().Select(label => label.GetString()).ToList(),
                Assignee = fields.GetProperty("assignee").GetProperty("displayName").GetString(),
                Reporter = fields.GetProperty("reporter").GetProperty("displayName").GetString(),
                Priority = fields.GetProperty("priority").GetProperty("name").GetString(),
                Status = fields.GetProperty("status").GetProperty("name").GetString(),
                Resolution = fields.GetProperty("resolution").GetProperty("name").GetString(),
                Created = DateTime.Parse(fields.GetProperty("created").GetString()),
                Updated = DateTime.Parse(fields.GetProperty("updated").GetString()),
                DueDate = fields.GetProperty("duedate").TryGetString(out var dueDateString) ? (DateTime?)DateTime.Parse(dueDateString) : null,
                Sprint = fields.GetProperty(sprintField).GetProperty("name").GetString(),
                EpicLink = fields.GetProperty("epic").GetProperty("key").GetString(),
                SubTasks = fields.GetProperty("subtasks").EnumerateArray().Select(subtask => subtask.GetProperty("key").GetString()).ToList(),
                Verrechenbarkeit = fields.TryGetProperty(chargableField)?.GetString()
            };


        }

    }

}
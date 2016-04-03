# open-search-result-links #
Specify a workbook (.xlsx) and the target column and this script will load each hyperlink in that column.

## To do ##
If going the route of opening every link, then inspecting results as a whole at the end:
- Halt for some time before opening the next link
If going the route of performing the task through CLI rather than in Microsoft Excel:
- Present the user with three questions after each link opens, with the response set to change the color of the corresponding row
    - Is the link safe?
        - If not, the link is not safe
    - Is the link from a customer site?
    - Is the link from an LMS?
        - This probably won't be needed, since it must be true if the responses to questiosn 1 and 2 are "No"
- Be able to change the color of each row depending on the response to those questions

# Project Roadmap & TODO

## New Features

- [ ] **Table View with Due Dates**
  - Implement a rendering mode that displays tasks in a Markdown table.
  - Columns should include: *Task Name*, *Due Date*, *Priority*, and *Status*.
  - Enable sorting options (e.g., sort by due date ascending).

- [ ] **Handling for deleted Tasks**
  - Implement handling for tasks that are deleted in ToDo
  - Likely remove them from Obsidian too, or place them in a hidden archived view

- [ ] **Page-Specific List Binding**
  - Allow binding one or multiple Microsoft To Do lists to a specific Obsidian page.
  - **Auto-Gathering:** Automatically fetch and append new tasks from the specific list to the page.
  - **Persistence:** Use a mechanism (e.g., a hidden subtask or frontmatter) to track the source list name/ID, ensuring tasks remain linked to their correct list even after completion.

- [ ] **Enhanced Collaboration Support**
  - Improve the handling of shared To Do lists.
  - Visualize `Assignee` and `Contributor` fields for tasks.
  - Ensure reliable syncing when multiple users modify the same shared list.

## User Interface Improvements

- [ ] **Visual Polish for Imported Tasks**
  - Improve the CSS styling of imported tasks to look more native and modern within Obsidian.
  - Consider adding icons or color-coding for priorities and tags.

- [ ] **Entra App Migration**
  - Register a new dedicated Azure (Entra ID) application for the plugin.
  - Move away from the original author's personal Client ID to ensure community control and long-term stability.

## Future Proofing

- [ ] **API & Framework Monitoring**
  - Implement an automated check (e.g., via GitHub Actions) to track changes in the Microsoft Graph API.
  - Monitor updates to the Obsidian Plugin API to catch deprecations early and alert maintainers when rework is required.

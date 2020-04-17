# README #

### Introduction

This little handy script allows to export a list of projects into OmniOutliner giving you a bird's eye view of your projects. Inspired by [Getting Things Done](https://gettingthingsdone.com/) methodology which references a project list - a physical sheet of paper with the names of your current projects.

>**Important note:** both, OmniFocus and OmniOutliner, require pro versions as I use AppleScript for both which is available only in Pro versions

### Features

- selective export:
- choose to export either active, on-hold or every project
- project entries in OmniOutliner are hyperlinks to OmniFocus projects
- in addition to `active` and `on-hold`, the following custom statuses are available
- `waiting`: all *available* tasks have been assigned `wait-for` tag (tag which has the word `wait` in its name)
- `stalled (no tags)`: all *remainining* tasks have no assigned tags
- `stalled (no tasks)`: there are no *remaining* tasks
- `deferred (tasks)`: there are no available tasks but there are remaining tasks
- `deferred (project)`: project has a defer date set to future date (when defer date lapses, it remains but it is referred to as expired)
- custom styles bound to function keys:
- `F1`: green highlight
- `F2`: orange highlight
- `F3`: red highlight
- there is a separate column called `Root Folder` which displays the root folder of the project include a sub-folder in parentheses*
- by default, projects are ordered by `Root Folder` in ascending manner

>**Important notice**: this script was designed with projects layout in mind where there are root level folders acting as logical grouping .e.g `work` or `home`; the sub-folder in the parentheses shows the name of the folder belonging to the second nesting level of folders

### OmniFocus 2

This script is meant to be used with OmniFocus 3 version. OmniFocus 3 introduced a breaking change where more than one `tag` (formerly known as `context`) could be assigned to a task. At some point, I have upgraded to OmniFocus 3 and tried to provide support for both versions. Unfortunately, it has proven to be too difficulty and I have decided to drop support for OmniFocus 2. The latest commit to support version 2 is https://github.com/alexkuc/omnifocus-projects-script/commit/d3289db5feddf9c520db3bd6436780d2e89d4ed9.

### OmniOutliner

I am currently on OmniOutliner 5 though technically script supports previous versions, 3 and 4. However, I do not have these versions locally installed so I cannot guarantee the experience will be bug free. In case you run into issue, please open [a new issue](https://github.com/alexkuc/omnifocus-projects-script/issues/new) so I can investigate and help you.

### How do I get set up? ###

* Copy script to OmniFocus scripts folder
* Drag script's icon to toolbar

>**Warning:** don't forget to set the encoding to `MacRoman` if you are going to edit the script

# Excel-VBA-Gradebook
Excel VBA Macro for Gradebooks

Current Version: 0.5

Current Features

- Detects the selected student's cell, and generates a progress report based on each of 5 categories (Assignments, Attendance & Participation, Tests, Midterm & Final Exams, and Semester Grade.

To Be Added

- ProgRepClass: Something that generates a progress report in each category for an entire class of students, regardless of the student or cell selected.
- Cell Resizing: Resizes the cell based on the assignment title length.
- Master Stats: A page in the gradebook that states the number of classes in the workbook, how many students in each class, and perhaps some other stats that can be useful at a glance.
- Column Split: I may take the first column-split feature and extend it to the other categories.

Hello all,

I'm a teacher in Guangzhou, China, and I usually store my gradebooks in excel files before I upload them to our LMS.

Our LMS is rather ungainly, made with Java, and has difficulty accessing the main server from China, making it rather inefficient for workaday tasks.

This project is designed to organize the different version of Excel macros that I write to process my students' grades into a progress report.

In time, I may decide to add other features and macros, but primarily this macro is designed to take a selection (of a student's name), read a number of variables, including the assignment categories selected, and how many assignments per category, and then generate a progress report.

Another version of the macro will generate progress reports for an entire class

Thus far, since I'm very new to Excel VBA, I have put everything into a single sub, without making much use of functions, though I have thought about how to resolve the various tasks into discrete functions, so as to cut down on the size of the script. Thus far, however, I just want something that practically works, rather than something that looks like what a professional coder does.

Nevertheless, I'm very open to feedback and criticism of the macro, and I'm obviously not looking to reinvent the wheel, but I haven't found anything like what I'm building. I imagine the reason for this is that most teachers today use an LMS required by their school, and live near enough to the server for performance to not be an issue. I have used Canvas before, but I"m not that big a fan. I do most of my work in OneNote, and building Excel attachments tends to be the most efficient way to manage my classes.

# patching-report-generator
Your manager wants you to fill excel files with 4895 rows every time your team patches some servers? Take this

This is essentially a parser with extra steps.
Use case:
1. server patching has been already automatised
2. patching software outputs csv logs after each session, these contain server names, whether servers have been patched successfully and additional comments
3. you need to fill a big excel sheet with all the statuses that auto patcher returned (because your manager loves big excel sheets)
4. you run this script (hopefully nothing breaks)
5. profit

I tried to make this as simple as possible and as quickly as possible. Not sure if simplicity was achieved but at least it works.

The script needs openpyxl to work and was developed under Linux. It will not run under MS Windows.

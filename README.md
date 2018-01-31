# ScrapReconciliation
Excel VBA code for reconciliation program 

The NUC_Raw_Materials_Reconciliation.xlsm.vba folder contains .bas files for all vba modules used in program.  

To push any code changes, clone this repo to your local machine.  There is a short python script (pre-commit.py) that you should run initially after cloning the repository.  This will automatically export your vba modules to .bas files when you commit changes to the remote repo.

The user form (userform1) is file type .frm.  If you make changes to the user form, you will have to manually export the new .frm file to your local branch before you commit.  The script only handles .bas files.  I'm looking into auto-exporting the .frm files as well.

The "Home" page is part of the NUC_Raw_Materials_Reconciliation.xlsm file in the repository.  If you make any changes to it, please communicate those changes either in your commit message or comment on the commit directly in GitHub.
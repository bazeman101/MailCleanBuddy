# How to Upload Your MailCleanBuddy Project to GitHub

This guide provides the steps to upload your MailCleanBuddy project (including `MailCleanBuddy.ps1`, `localizations.json`, and `README.md`) to a new GitHub repository.

**Prerequisites:**

*   **Git Installed:** Ensure you have Git installed on your system. You can download it from [git-scm.com](https://git-scm.com/).
*   **GitHub Account:** You need a GitHub account. If you don't have one, sign up at [github.com](https://github.com/).

**Steps:**

1.  **Navigate to Your Project Directory:**
    Open your terminal (like PowerShell, Command Prompt, Git Bash, or your macOS/Linux terminal).
    Use the `cd` (change directory) command to go to the folder where your project files (`MailCleanBuddy.ps1`, `localizations.json`, `README.md`) are located.
    ```bash
    # Example:
    # cd C:\_work\OutlookBuddy 
    # or
    # cd /path/to/your/project
    ```

2.  **Initialize a Local Git Repository:**
    If your project folder is not already a Git repository (e.g., if you don't see a `.git` subfolder), you need to initialize it. If you've used Git for this project before and a `.git` folder exists, you can skip this step.
    ```bash
    git init
    ```

3.  **Check Git Status (Optional but Recommended):**
    See which files Git is aware of and their status.
    ```bash
    git status
    ```
    If you just initialized, you should see your project files listed as untracked files.

4.  **Add Your Files to Git's Staging Area:**
    This tells Git which files you want to include in your next commit.
    To add all relevant files:
    ```bash
    git add MailCleanBuddy.ps1 localizations.json README.md TOGITHUB.md
    ```
    Alternatively, to add all files and subdirectories in the current folder:
    ```bash
    git add .
    ```

5.  **Commit Your Files:**
    This saves a snapshot of your staged files to your local Git repository's history.
    ```bash
    git commit -m "Initial commit of MailCleanBuddy project"
    ```
    You can replace `"Initial commit of MailCleanBuddy project"` with any descriptive message.

6.  **Create a New Repository on GitHub:**
    *   Go to [GitHub](https://github.com/) and log in.
    *   In the top-right corner, click the **+** icon, then select **New repository**.
    *   **Repository name:** Choose a name for your repository (e.g., `MailCleanBuddy` or `MailCleanBuddy-PS`).
    *   **Description:** (Optional) Add a brief description of your project.
    *   **Public/Private:** Choose whether you want your repository to be public (visible to everyone) or private (visible only to you and collaborators you choose).
    *   **Important:** **Do NOT** initialize the repository with a `README`, `.gitignore`, or `license` if you have already created these files locally (you have `README.md`). Uncheck these options if they are selected.
    *   Click the **Create repository** button.

7.  **Link Your Local Repository to the GitHub Repository:**
    After creating the repository on GitHub, you'll be taken to its main page. GitHub will display instructions, including a URL for your new repository. It will look like `https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git`.
    Copy this URL.

    Back in your terminal, in your project directory, run the following command, replacing `YOUR_USERNAME` and `YOUR_REPOSITORY_NAME` with your actual GitHub username and the repository name you chose:
    ```bash
    git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git
    ```
    This command tells your local Git repository that there's a remote repository named `origin` at the specified URL. If you previously had an `origin` remote, you might need to remove it first (`git remote remove origin`) or update it (`git remote set-url origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git`).

8.  **Set Up Authentication (Using a Personal Access Token - PAT):**
    GitHub no longer supports password authentication for Git operations over HTTPS. You need to use a Personal Access Token (PAT).

    *   **Create a PAT on GitHub:**
        1.  Go to your GitHub settings: Click your profile picture in the top-right corner, then click **Settings**.
        2.  In the left sidebar, scroll down and click **Developer settings**.
        3.  In the left sidebar, click **Personal access tokens**, then **Tokens (classic)**.
        4.  Click **Generate new token**, then **Generate new token (classic)**.
        5.  Give your token a descriptive **Note** (e.g., "MailCleanBuddy CLI Access").
        6.  Set an **Expiration** period for your token (e.g., 30 days, 90 days, or custom). For security, tokens should expire.
        7.  Under **Select scopes**, check the box next to **`repo`**. This will grant the token permissions to access your repositories (public and private), including pushing code.
        8.  Scroll down and click **Generate token**.
        9.  **Important:** GitHub will show you your new PAT **only once**. Copy it immediately and save it in a secure place (like a password manager). You will not be able to see it again. If you lose it, you'll have to generate a new one.

    *   **Using the PAT:** When Git prompts you for a password (e.g., during `git push`), enter your PAT instead of your GitHub account password. Your username will still be your GitHub username.

9.  **Verify the Remote Repository (Optional):**
    To ensure the remote was added correctly:
    ```bash
    git remote -v
    ```
    You should see output showing `origin` followed by your GitHub repository URL for both fetch and push.

10. **Push Your Local Commits to GitHub:**
    This uploads your committed files and their history from your local repository to the remote repository on GitHub.

    *   **Determine your default branch name:** Git's default branch name used to be `master`, but newer versions often use `main`. You can check your current branch name by running `git branch`. The one with an asterisk `*` next to it is your current branch.
    *   If your branch is `main`:
        ```bash
        git push -u origin main
        ```
    *   If your branch is `master`:
        ```bash
        git push -u origin master
        ```

    The `-u` flag (short for `--set-upstream`) links your local branch to the remote branch. After the first push with `-u`, you can simply use `git push` for subsequent pushes from that branch.

    *   **Authentication:** When prompted for your username, enter your GitHub username. When prompted for your password, **paste your Personal Access Token (PAT)** that you created in step 8. Some Git credential managers might cache this for you after the first successful authentication.

11. **Verify on GitHub:**
    Refresh your repository page on GitHub. You should now see your `MailCleanBuddy.ps1`, `localizations.json`, `README.md`, and `TOGITHUB.md` files listed.

Your project is now on GitHub!

---
**Alternative: SSH Authentication**
Instead of HTTPS with a PAT, you can use SSH keys for authentication. This involves generating an SSH key pair, adding the public key to your GitHub account, and changing your remote URL to the SSH format (e.g., `git@github.com:YOUR_USERNAME/YOUR_REPOSITORY_NAME.git`). This method is often preferred for frequent use as it doesn't require entering a token each time (if your SSH key is managed by an SSH agent). You can find more details on GitHub's documentation for [Connecting to GitHub with SSH](https://docs.github.com/en/authentication/connecting-to-github-with-ssh).

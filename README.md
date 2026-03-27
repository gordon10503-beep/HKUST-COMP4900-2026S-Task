# HKUST 2026 Spring COMP 4900 GitHub Training Module Task

Welcome to the COMP 4900 GitHub Training Module task. You can review the training module material website [here](https://minglol7794.github.io/github_tutorial/COMP4900-Website)

## Task

In this exercise, you will practice a standard open-source contribution workflow:
1. Fork a repository
2. Clone your fork locally
3. Create a branch
4. Edit your student record
5. Push your branch
6. Open a pull request

For a detailed explanation of the training task, please refer to:
[COMP4900 GitHub Training Website](https://minglol7794.github.io/github_tutorial/COMP4900-Website/#task-instructions)

After your pull request is validated, you will submit the pull request link on Canvas.

## Before You Start

Please make sure you have:
1. A GitHub account
2. Git installed on your computer

## Step-by-Step Instructions

## Step 1: Fork This Repository

Click the **Fork** button on GitHub to create your own copy of this repository under your account.

## Step 2: Clone Your Fork

Clone your forked repository to your local machine:

```bash
git clone <your-fork-url>
cd COMP4900
```

Example fork URL:

```bash
https://github.com/<your-github-username>/pull-request-test.git
```

## Step 3: Create a New Branch and Edit Data/student.json

Create a new branch before making changes:

```bash
git checkout -b add-my-student-info
```

Open `Data/student.json` and add or update your record with:
1. Name
2. Student ID
3. ITSC email
4. GitHub username

Please keep the JSON format valid (do not break commas, brackets, or quotes).

## Step 4: Commit and Push Your Changes

Stage, commit, and push your branch:

```bash
git add Data/student.json
git commit -m "Add my COMP4900 student information"
git push origin add-my-student-info
```

## Step 5: Open a Pull Request

On GitHub, open a pull request from your branch in your fork to the target branch of the original repository.

Make sure your pull request clearly shows your student information update.

## Step 6: Wait for Auto-Validation

If your submission is correct:
1. You will receive an automatic reply in the pull request
![Success submit automatic reply](Assets/task-success-submit.png)
2. Your record will be updated in the tracking spreadsheet

You can check whether your pull request is successfully recorded here:
[Pull Request Tracking Spreadsheet](https://docs.google.com/spreadsheets/d/1P81-TPUOiYIHPcdzJMZH5RLOhV9ChRGGtxq_yakSkxw/edit?gid=0#gid=0)

If validation fails, read the bot message, fix the issue, and push updates to the same branch.

## Step 7: Submit on Canvas

Submit your pull request link to Canvas:

[Canvas submission link](https://canvas.ust.hk/courses/67791/assignments/424723?module_item_id=1753451)

## Quick Checklist

Before submitting, confirm that:
1. You forked the repository
2. You created a separate branch
3. You only edited `Data/student.json` for your info update
4. Your JSON format is valid
5. Your pull request is open and visible
6. You submitted the pull request link on Canvas

## Extra
Here is a Nessie, cute.
![Nessie](Assets/NESSIE.png)

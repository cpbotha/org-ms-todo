;; org-ms-todo.el -- sync your org-mode tasks with Microsoft To-Do

;; Copyright (C) 2024 Charl P. Botha

;; Author: Charl P. Botha <cpbotha@vxlabs.com>
;; Version: 1.0
;; Package-Requires: ((oauth2-auto "xxx")(org-ql "xxx"))
;; Keywords: org, tasks, microsoft, todo
;; URL: https://github.com/cpbotha/org-ms-todo

;; TODO:
;; - [X] when org task is DONE, update MS task to DONE
;; - [X] when MS task is DONE, update org task to DONE
;; - [X] Consider also setting the body of the MS to-do
;; - [X] Consider setting tags via ms todo hashtags
;; - [X] Bundle the instructions at the end into org-ms-todo-sync function
;; - [X] Only do HTML export when that task is actually going to be created on the MS side
;; - [ ] Retrieve MS list id based on the list name "Emacs Org" (conifgurable)
;; - [ ] Consider creating new org-mode tasks from MS to-dos in the correct group

;; https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http
;; https://learn.microsoft.com/en-us/graph/api/todotasklist-list-tasks?view=graph-rest-1.0&tabs=http

;; if you have PGP setup also for use in Emacs, configure plstore-encrypt-to
;; for public key encryption instead of symmetric. With symmetric, you're going
;; to have to type the plstore passphrase more often than you like.

;; temporary setup for testing ---------------------------------------------
(use-package oauth2-auto
  :ensure t
  ;; with just the git repo https://github.com/telotortium/emacs-oauth2-auto.git this would fail 
  :vc (:fetcher "https://raw.githubusercontent.com/telotortium/emacs-oauth2-auto/main/oauth2-auto.el")
  )

(use-package org-ql
  :ensure t)

;; so that plstore will use pgp and not symmetric
;; TODO: move to init
(setq plstore-encrypt-to "0xE77A4564")

(setq emacs-list-id  "AQMkADAwATM3ZmYAZS0wMjRiLTg5YTQtMDACLTAwCgAuAAADJ6bSE4UnyEaAs4sDMsoQdgEAd96ZmABYYkKEE5goOL-SRQAGhuyhVAAAAA==")

;; end of temporary setup ---------------------------------------------------



(require 'oauth2-auto)
(require 'org-ql)
(require 'request)

(setq ms-oauth2-url (concat "https://login.microsoftonline.com/"
                            oauth2-auto-microsoft-default-tenant
                            "/oauth2/v2.0/")


      ;; cpbotha-org-ms-todo application (client id)
      ;; as far as I can tell, I can distribute this, and other folks can use this code to sync their tasks via my app
      oauth2-auto-microsoft-client-id "a1a354c4-babe-404b-9207-101d0b836847"


      )

(setq oauth2-auto-additional-providers-alist
      `((microsoft-todo
         (authorize_url . ,(concat ms-oauth2-url "authorize"))
         (token_url . ,(concat ms-oauth2-url "token"))
         (tenant . ,oauth2-auto-microsoft-default-tenant)
         (scope . "offline_access Tasks.ReadWrite")
         (client_id . ,oauth2-auto-microsoft-client-id)
         )))


;; if you also want the refresh-token, use the following instead:
;;(setq ms-auth (oauth2-auto-plist-sync "cpbotha@vxlabs.com" 'microsoft-todo))
;; get :access-token from ms-auth
;;(setq access-token (plist-get ms-auth :access-token))


;; example from the docs to create new task:
;; POST https://graph.microsoft.com/v1.0/me/todo/lists/AQMkADAwATM0MDAAMS0yMDkyLWVjMzYtM/tasks
;; Content-Type: application/json
;; {
;;    "title":"A new task",
;;    "categories": ["Important"],
;; }



(defun org-ms-todo--ms-list-task-lists ()
  (request
    "https://graph.microsoft.com/v1.0/me/todo/lists"
    :type "GET"
    :headers `(("Authorization" . ,(format "Bearer %s" access-token)))
    :parser 'json-read
    :success (cl-function
              (lambda (&key data &allow-other-keys)
                (message "Got task lists: %S" data)))))

;; test the list-tasks-lists function 
;; use this to get the ID of the list you want to use, in my case "emacs ..."
;;(org-ms-todo--ms-list-task-lists)


;; create a new task inside the emacs list
(defun org-ms-todo--ms-create-task (ms-list-id title body-html org-id due-datetime scheduled-datetime timezone)
  (setq create-data-json (json-encode
                          ;; use append so that optional fields can be nil, and then are completely excluded from the alist
                          (append
                           `((title . ,title))
                           ;; :dueDateTime (:dateTime 2024-04-01T22:00:00.0000000 :timeZone UTC) 
                           (when due-datetime
                             `((dueDateTime . ((dateTime . ,due-datetime) (timeZone . ,timezone )))))
                           (when scheduled-datetime
                             `((startDateTime . ((dateTime . ,scheduled-datetime) (timeZone . ,timezone )))))
                           (when body-html `((body . ((contentType . "html") (content . ,body-html)))))
                           `(("linkedResources" . ,(list `(("applicationName" . "Emacs Orgmode") 
                                                           ;; add org-protocol link here as well, because to-do web-app does not link or even reveal the webUrl
                                                           ("displayName" . ,(format "org-protocol://org-id?id=%s" org-id))
                                                           ("externalId" . ,org-id)
                                                           ("webUrl" . ,(format "org-protocol://org-id?id=%s" org-id))
                                                           ))))
                           nil)))
  (message "Create task with: %s" create-data-json)
  (request
    (format "https://graph.microsoft.com/v1.0/me/todo/lists/%s/tasks" ms-list-id)
    :type "POST"
    :headers `(("Content-Type" . "application/json") 
               ("Authorization" . ,(format "Bearer %s" access-token))
               ;; prefer header required, else HTML stripped from html body
               ("Prefer". "outlook.body-content-type='html'"))
    :data create-data-json
    :parser 'json-read
    :success (cl-function
              (lambda (&key data &allow-other-keys)
                (message "Created task: %S" data)))))

;; https://learn.microsoft.com/en-us/graph/api/todotask-update?view=graph-rest-1.0&tabs=http
;; status: notStarted, inProgress, completed, waitingOnOthers, deferred
(defun org-ms-todo--ms-update-task (ms-list-id ms-task-id status completed-datetime timezone)
  (setq update-data-json (json-encode
                          (append
                           `((status . ,status))
                           ;; :completedDateTime (:dateTime 2024-04-01T22:00:00.0000000 :timeZone UTC) 
                           (when completed-datetime `((completedDateTime . ((dateTime . ,completed-datetime) (timeZone . ,timezone )))))
                           nil)))
  (message "Update task with: %s" create-data-json)
  (request
    (format "https://graph.microsoft.com/v1.0/me/todo/lists/%s/tasks/%s" ms-list-id ms-task-id)
    :type "PATCH"
    :headers `(("Content-Type" . "application/json") 
               ("Authorization" . ,(format "Bearer %s" access-token)))
    :data update-data-json
    :parser 'json-read
    :success (cl-function
              (lambda (&key data &allow-other-keys)
                (message "Updated task: %S" data)))))

(defun org-ms-todo--ms-list-tasks (list-id)
  (request
    (format "https://graph.microsoft.com/v1.0/me/todo/lists/%s/tasks" list-id)
    :type "GET"
    :sync t
    :headers `(("Authorization" . ,(format "Bearer %s" access-token))
               ("Prefer". "outlook.body-content-type='html'"))
    ;; json-read by default gives alists
    ;; here we wrap it to give plist, as per docs
    :parser (lambda ()
              (let ((json-object-type 'plist))
                (json-read)))
    :success (cl-function
              (lambda (&key data &allow-other-keys)
                (setq ms-tasks (append (plist-get data :value) nil))
                ;; DEBUG
                ;; (message "Got task lists: %S" data)
                ))))

;; example using org-ql-query
;;https://github.com/alphapapa/org-ql/blob/master/examples.org#listing-bills-coming-due
;; org-ql-select could also be used for this
;; https://github.com/alphapapa/org-ql/tree/master?tab=readme-ov-file#function-org-ql-select

;; NOTE: org-ql can search via the "CLOSED:" property

;; algo
;; - get all org tasks with ID property
;; - for each task, ensure there's a corresponding MS task. If not, create.

;; convert org-element timestamp of the form:
;; (timestamp (:type active :raw-value <2024-03-22 Fri> :year-start 2024 :month-start 3 :day-start 22 :hour-start nil :minute-start nil :year-end 2024 :month-end 3 :day-end 22 :hour-end nil :minute-end nil :begin 12111 :end 12127 :post-blank 0))
;; ... into iso-format timestamp, in the UTC timezone
(defun org-ms-todo--org-timestamp-to-iso (org-timestamp)
  "Convert ORG-TIMESTAMP to iso-format timestamp.

If ORG-TIMESTAMP is nil, return nil. "
  (when org-timestamp (format-time-string "%Y-%m-%dT%H:%M:%S" (org-timestamp-to-time org-timestamp))))


;; this task is an org-element
(defun org-ms-todo--handle-org-task (task-oe)
  (let* ((task (car (cdr task-oe)))
         (id (plist-get task :ID))
         (title (plist-get task :raw-value))
         (ms-task (seq-find (lambda (tsk)
                              ;; does plist-get tsk :linkedResources contain :externalId == id
                              (seq-find (lambda (lr) (string= (plist-get lr :externalId) id))
                                        (plist-get tsk :linkedResources))
                              ) ms-tasks)))
    (if ms-task
        (progn
          (message "Task with org ID %s already exists in MS TODO" id)
          ;; if either side is DONE, update the other side if not already DONE

          ;; org-task :todo-type will be 'todo or 'done, user-defined keywords automatically grouped
          ;; if org-task is DONE, but ms-task is not, update ms-task
          (if (eq (plist-get task :todo-type) 'done)
              (unless (string= (plist-get ms-task :status) "completed")
                (message "Update MS to-do with org ID %s to completed" id)
                (org-ms-todo--ms-update-task emacs-list-id (plist-get ms-task :id) 'completed (org-ms-todo--org-timestamp-to-iso (plist-get task :closed)) "Africa/Johannesburg"))
            ;; org-task is in todo state; if ms-task is done, update org-task to done
            (when (string= (plist-get ms-task :status) "completed")
              ;; change org-task to done
              (message "Queueing update of org task with ID %s to DONE" id)
              ;; convert completedDateTime dateTime (iso formatted) to org timestamp
              ;; :completedDateTime (:dateTime "2024-03-21T22:00:00.0000000" :timeZone "UTC")
              ;; FIXME: we assume that :timeZone is UTC, and as if that was not bad enough, we actually just ignore it and use the time as is
              ;; store both org-id and completed timestamp (as lisp timestamp) in the queue
              (push `(,id ,(parse-iso8601-time-string (plist-get (plist-get ms-task :completedDateTime) :dateTime))) org-ms-todo--queue-done))))

      (progn
        ;; we have org-task, but no corresponding ms-task
        ;; only when the org-task is still in todo state do we create the ms to-do
        (when (eq (plist-get task :todo-type) 'todo)
          ;; according to https://orgmode.org/worg/dev/org-element-api.html
          ;; :tags should be a list of strings, but I have to do this extra dance
          ;; to get out just the tag strings
          (let ((hashtags (mapconcat (lambda (tag)
                                       (concat "#" (substring-no-properties tag)))
                                     (org-element-property :tags (car org-tasks)) " "))
                (body-html (with-current-buffer (plist-get task :element-buffer)
                             (org-ms-todo--body-to-html task-oe))))
            (org-ms-todo--ms-create-task emacs-list-id
                                         (concat title (when hashtags " ") hashtags) 
                                         body-html
                                         id
                                         (org-ms-todo--org-timestamp-to-iso (plist-get task :deadline))
                                         (org-ms-todo--org-timestamp-to-iso (plist-get task :scheduled))
                                         "Africa/Johannesburg")
            (message "Create new task with title %s" title))))

      )))

;; heading is an org-element (heading (:raw-value ...))
;; this function will only work during org-ql-select, when the relevant buffer is current!
(defun org-ms-todo--body-to-html (heading)
  "Extract body of org HEADING element and convert to HTML."
  ;; robust-* skips the scheduling line and the properties drawer
  ;; logbook will be included however
  (let* ((beg (org-element-property :robust-begin heading))
         (end (org-element-property :robust-end heading))
         (body (when (and beg end) (buffer-substring-no-properties beg end)))
         (body-html (when body (org-export-string-as
                                body
                                'html t '(:with-toc nil :with-tags nil)))))
    body-html))


;; if you have many TODO tasks that are going to sync, this could give you 429 errors
;; just re-run
(defun org-ms-todo-sync()

  (interactive)

  ;; this will give you just the access-token 
  (setq access-token (oauth2-auto-access-token-sync "cpbotha@vxlabs.com" 'microsoft-todo))
  
  ;; this returns a list of alists
  ;; note that we're in sync mode, so after this we have the data
  ;; FIXME: pass ms-tasks in as a parameter
  (setq ms-tasks nil)
  (org-ms-todo--ms-list-tasks emacs-list-id)

  ;; get list of org agenda TODOs and DONEs
  ;; https://github.com/alphapapa/org-ql/blob/master/examples.org
  ;; org-ql-search is interactive

  ;; when we parse tasks, we also store each task's buffer in the task itself
  ;; we might need this later if we need to access the task's body (for HTML export)
  (setq org-tasks
        (org-ql-select
          (org-agenda-files)
          '(and (property "ID") (or (todo) (done)))
          :action
          '(let* ((hl (org-element-headline-parser)))
             (org-element-put-property hl :element-buffer (current-buffer))
             hl)))

  ;; FIXME: pass queue in as a parameter
  (setq org-ms-todo--queue-done nil)
  (mapc (lambda (org-task) (org-ms-todo--handle-org-task org-task)) org-tasks)
  ;; DEBUG: only do the first one
  ;; (org-ms-todo--handle-org-task (car org-tasks))

  ;; now we have a list of org tasks that need to be updated to DONE
  ;; https://github.com/alphapapa/org-ql/blob/master/examples.org#set-tags-on-certain-entries
  (mapc (lambda (id-ts)
          (org-ql-select (org-agenda-files) '(and (property "ID")
                                                  (or (todo) (done))
                                                  (string-match  (car id-ts) (org-entry-get (point) "ID")))
            :action '(progn
                       ;; toggle to DONE
                       (org-todo 'done)
                       ;; set CLOSED special property to what we received from ms todo
                       (org-add-planning-info 'closed (car (cdr id-ts))))))

        org-ms-todo--queue-done)



  )

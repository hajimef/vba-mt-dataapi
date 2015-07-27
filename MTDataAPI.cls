VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MTDataAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private baseUrl As String
Private clId As String
Private accessToken As String
Private jlib As JSONLib
Private js As ScriptControl
Private fso As Scripting.FileSystemObject
Private methods As Scripting.Dictionary

' コンストラクタ
Private Sub Class_Initialize()
    Dim method As Scripting.Dictionary, resources As Collection
    Set jlib = New JSONLib
    Set js = New ScriptControl
    Set fso = New Scripting.FileSystemObject
    js.Language = "Jscript"
    Set methods = New Scripting.Dictionary
    ' listEndpoints
    Set method = New Scripting.Dictionary
    method("route") = "/endpoints"
    method("verb") = "GET"
    Set methods("listEndpoints") = method
    ' authorize
    Set method = New Scripting.Dictionary
    method("route") = "/authorization"
    method("verb") = "GET"
    Set methods("authorize") = method
    ' authenticate
    Set method = New Scripting.Dictionary
    method("route") = "/authentication"
    method("verb") = "POST"
    Set methods("authenticate") = method
    ' getToken
    Set method = New Scripting.Dictionary
    method("route") = "/token"
    method("verb") = "POST"
    Set methods("getToken") = method
    ' revokeAuthentication
    Set method = New Scripting.Dictionary
    method("route") = "/authentication"
    method("verb") = "DELETE"
    Set methods("revokeAuthentication") = method
    ' revokeToken
    Set method = New Scripting.Dictionary
    method("route") = "/token"
    method("verb") = "DELETE"
    Set methods("revokeToken") = method
    ' getUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id"
    method("verb") = "GET"
    Set methods("getUser") = method
    ' updateUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "user"
    Set method("resources") = resources
    Set methods("updateUser") = method
    ' listBlogsForUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id/sites"
    method("verb") = "GET"
    Set methods("listBlogsForUser") = method
    ' getBlog
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id"
    method("verb") = "GET"
    Set methods("getBlog") = method
    ' listEntries
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries"
    method("verb") = "GET"
    Set methods("listEntries") = method
    ' createEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "entry"
    Set method("resources") = resources
    Set methods("createEntry") = method
    ' getEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id"
    method("verb") = "GET"
    Set methods("getEntry") = method
    ' updateEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "entry"
    Set method("resources") = resources
    Set methods("updateEntry") = method
    ' deleteEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id"
    method("verb") = "DELETE"
    Set methods("deleteEntry") = method
    ' previewEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/preview"
    method("verb") = "POST"
    Set methods("previewEntry") = method
    ' previewEntryById
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id/preview"
    method("verb") = "POST"
    Set methods("previewEntryById") = method
    ' listCategories
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories"
    method("verb") = "GET"
    Set methods("listCategories") = method
    ' listComments
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/comments"
    method("verb") = "GET"
    Set methods("listComments") = method
    ' listCommentsForEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id/comments"
    method("verb") = "GET"
    Set methods("listCommentsForEntry") = method
    ' createComment
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id/comments"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "comment"
    Set method("resources") = resources
    Set methods("createComment") = method
    ' createReplyComment
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id/comments/:comment_id/replies"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "comment"
    Set method("resources") = resources
    Set methods("createReplyComment") = method
    ' getComment
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/comments/:comment_id"
    method("verb") = "GET"
    Set methods("getComment") = method
    ' updateComment
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/comments/:comment_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "comment"
    Set method("resources") = resources
    Set methods("updateComment") = method
    ' deleteComment
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/comments/:comment_id"
    method("verb") = "DELETE"
    Set methods("deleteComment") = method
    ' listTrackbacks
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/trackbacks"
    method("verb") = "GET"
    Set methods("listTrackbacks") = method
    ' listTrackbacksForEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id/trackbacks"
    method("verb") = "GET"
    Set methods("listTrackbacksForEntry") = method
    ' getTrackback
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/trackbacks/:ping_id"
    method("verb") = "GET"
    Set methods("getTrackback") = method
    ' updateTrackback
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/trackbacks/:ping_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "trackback"
    Set method("resources") = resources
    Set methods("updateTrackback") = method
    ' deleteTrackback
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/trackbacks/:ping_id"
    method("verb") = "DELETE"
    Set methods("deleteTrackback") = method
    ' uploadAsset
    Set method = New Scripting.Dictionary
    method("route") = "/assets/upload"
    method("verb") = "POST"
    Set methods("uploadAsset") = method
    ' listPermissionsForUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id/permissions"
    method("verb") = "GET"
    Set methods("listPermissionsForUser") = method
    ' publishEntries
    Set method = New Scripting.Dictionary
    method("route") = "/publish/entries"
    method("verb") = "GET"
    Set methods("publishEntries") = method
    ' getStatsProvider
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/stats/provider"
    method("verb") = "GET"
    Set methods("getStatsProvider") = method
    ' listStatsPageviewsForPath
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/stats/path/pageviews"
    method("verb") = "GET"
    Set methods("listStatsPageviewsForPath") = method
    ' listStatsVisitsForPath
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/stats/path/visits"
    method("verb") = "GET"
    Set methods("listStatsVisitsForPath") = method
    ' listStatsPageviewsForDate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/stats/date/pageviews"
    method("verb") = "GET"
    Set methods("listStatsPageviewsForDate") = method
    ' listStatsVisitsForDate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/stats/date/visits"
    method("verb") = "GET"
    Set methods("listStatsVisitsForDate") = method
    ' listCategoriesForEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id/categories"
    method("verb") = "GET"
    Set methods("listCategoriesForEntry") = method
    ' listParentCategories
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories/:category_id/parents"
    method("verb") = "GET"
    Set methods("listParentCategories") = method
    ' listSiblingCategories
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories/:category_id/siblings"
    method("verb") = "GET"
    Set methods("listSiblingCategories") = method
    ' listChildCategories
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories/:category_id/children"
    method("verb") = "GET"
    Set methods("listChildCategories") = method
    ' createCategory
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "category"
    Set method("resources") = resources
    Set methods("createCategory") = method
    ' getCategory
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories/:category_id"
    method("verb") = "GET"
    Set methods("getCategory") = method
    ' updateCategory
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories/:category_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "category"
    Set method("resources") = resources
    Set methods("updateCategory") = method
    ' deleteCategory
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories/:category_id"
    method("verb") = "DELETE"
    Set methods("deleteCategory") = method
    ' permutateCategories
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories/permutate"
    method("verb") = "POST"
    Set methods("permutateCategories") = method
    ' listFolders
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders"
    method("verb") = "GET"
    Set methods("listFolders") = method
    ' listParentFolders
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders/:folder_id/parents"
    method("verb") = "GET"
    Set methods("listParentFolders") = method
    ' listSiblingFolders
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders/:folder_id/siblings"
    method("verb") = "GET"
    Set methods("listSiblingFolders") = method
    ' listChildFolders
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders/:folder_id/children"
    method("verb") = "GET"
    Set methods("listChildFolders") = method
    ' createFolder
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "folder"
    Set method("resources") = resources
    Set methods("createFolder") = method
    ' getFolder
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders/:folder_id"
    method("verb") = "GET"
    Set methods("getFolder") = method
    ' updateFolder
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders/:folder_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "folder"
    Set method("resources") = resources
    Set methods("updateFolder") = method
    ' deleteFolder
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders/:folder_id"
    method("verb") = "DELETE"
    Set methods("deleteFolder") = method
    ' permutateFolders
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders/permutate"
    method("verb") = "POST"
    Set methods("permutateFolders") = method
    ' listAssets
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/assets"
    method("verb") = "GET"
    Set methods("listAssets") = method
    ' listAssetsForEntry
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/:entry_id/assets"
    method("verb") = "GET"
    Set methods("listAssetsForEntry") = method
    ' listAssetsForPage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id/assets"
    method("verb") = "GET"
    Set methods("listAssetsForPage") = method
    ' listAssetsForSiteAndTag
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/tags/:tag_id/assets"
    method("verb") = "GET"
    Set methods("listAssetsForSiteAndTag") = method
    ' uploadAssetForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/assets/upload"
    method("verb") = "POST"
    Set methods("uploadAssetForSite") = method
    ' getAsset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/assets/:asset_id"
    method("verb") = "GET"
    Set methods("getAsset") = method
    ' updateAsset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/assets/:asset_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "asset"
    Set method("resources") = resources
    Set methods("updateAsset") = method
    ' deleteAsset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/assets/:asset_id"
    method("verb") = "DELETE"
    Set methods("deleteAsset") = method
    ' getThumbnail
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/assets/:asset_id/thumbnail"
    method("verb") = "GET"
    Set methods("getThumbnail") = method
    ' listEntriesForCategory
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/categories/:category_id/entries"
    method("verb") = "GET"
    Set methods("listEntriesForCategory") = method
    ' listEntriesForAsset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/assets/:asset_id/entries"
    method("verb") = "GET"
    Set methods("listEntriesForAsset") = method
    ' listEntriesForSiteAndTag
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/tags/:tag_id/entries"
    method("verb") = "GET"
    Set methods("listEntriesForSiteAndTag") = method
    ' importEntries
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/import"
    method("verb") = "POST"
    Set methods("importEntries") = method
    ' exportEntries
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/entries/export"
    method("verb") = "GET"
    Set methods("exportEntries") = method
    ' listPages
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages"
    method("verb") = "GET"
    Set methods("listPages") = method
    ' listPagesForFolder
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/folders/:folder_id/pages"
    method("verb") = "GET"
    Set methods("listPagesForFolder") = method
    ' listPagesForAsset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/assets/:asset_id/pages"
    method("verb") = "GET"
    Set methods("listPagesForAsset") = method
    ' listPagesForSiteAndTag
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/tags/:tag_id/pages"
    method("verb") = "GET"
    Set methods("listPagesForSiteAndTag") = method
    ' createPage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "page"
    Set method("resources") = resources
    Set methods("createPage") = method
    ' getPage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id"
    method("verb") = "GET"
    Set methods("getPage") = method
    ' updatePage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "page"
    Set method("resources") = resources
    Set methods("updatePage") = method
    ' deletePage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id"
    method("verb") = "DELETE"
    Set methods("deletePage") = method
    ' listCommentsForPage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id/comments"
    method("verb") = "GET"
    Set methods("listCommentsForPage") = method
    ' createCommentForPage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id/comments"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "comment"
    Set method("resources") = resources
    Set methods("createCommentForPage") = method
    ' createReplyCommentForPage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id/comments/:comment_id/replies"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "comment"
    Set method("resources") = resources
    Set methods("createReplyCommentForPage") = method
    ' listTrackbacksForPage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id/trackbacks"
    method("verb") = "GET"
    Set methods("listTrackbacksForPage") = method
    ' previewPage
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/preview"
    method("verb") = "POST"
    Set methods("previewPage") = method
    ' previewPageById
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/pages/:page_id/preview"
    method("verb") = "POST"
    Set methods("previewPageById") = method
    ' listSites
    Set method = New Scripting.Dictionary
    method("route") = "/sites"
    method("verb") = "GET"
    Set methods("listSites") = method
    ' listSitesByParent
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/children"
    method("verb") = "GET"
    Set methods("listSitesByParent") = method
    ' insertNewBlog
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "blog"
    Set method("resources") = resources
    Set methods("insertNewBlog") = method
    ' insertNewWebsite
    Set method = New Scripting.Dictionary
    method("route") = "/sites"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "website"
    Set method("resources") = resources
    Set methods("insertNewWebsite") = method
    ' updateSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id"
    method("verb") = "PUT"
    Set methods("updateSite") = method
    ' deleteSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id"
    method("verb") = "DELETE"
    Set methods("deleteSite") = method
    ' listRoles
    Set method = New Scripting.Dictionary
    method("route") = "/roles"
    method("verb") = "GET"
    Set methods("listRoles") = method
    ' createRole
    Set method = New Scripting.Dictionary
    method("route") = "/roles"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "role"
    Set method("resources") = resources
    Set methods("createRole") = method
    ' getRole
    Set method = New Scripting.Dictionary
    method("route") = "/roles/:role_id"
    method("verb") = "GET"
    Set methods("getRole") = method
    ' updateRole
    Set method = New Scripting.Dictionary
    method("route") = "/roles/:role_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "role"
    Set method("resources") = resources
    Set methods("updateRole") = method
    ' deleteRole
    Set method = New Scripting.Dictionary
    method("route") = "/roles/:role_id"
    method("verb") = "DELETE"
    Set methods("deleteRole") = method
    ' listPermissions
    Set method = New Scripting.Dictionary
    method("route") = "/permissions"
    method("verb") = "GET"
    Set methods("listPermissions") = method
    ' listPermissionsForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/permissions"
    method("verb") = "GET"
    Set methods("listPermissionsForSite") = method
    ' listPermissionsForRole
    Set method = New Scripting.Dictionary
    method("route") = "/roles/:role_id/permissions"
    method("verb") = "GET"
    Set methods("listPermissionsForRole") = method
    ' grantPermissionToSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/permissions/grant"
    method("verb") = "POST"
    Set methods("grantPermissionToSite") = method
    ' grantPermissionToUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id/permissions/grant"
    method("verb") = "POST"
    Set methods("grantPermissionToUser") = method
    ' revokePermissionFromSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/permissions/revoke"
    method("verb") = "POST"
    Set methods("revokePermissionFromSite") = method
    ' revokePermissionFromUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id/permissions/revoke"
    method("verb") = "POST"
    Set methods("revokePermissionFromUser") = method
    ' search
    Set method = New Scripting.Dictionary
    method("route") = "/search"
    method("verb") = "GET"
    Set methods("search") = method
    ' listLogs
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/logs"
    method("verb") = "GET"
    Set methods("listLogs") = method
    ' getLog
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/logs/:log_id"
    method("verb") = "GET"
    Set methods("getLog") = method
    ' createLog
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/logs"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "log"
    Set method("resources") = resources
    Set methods("createLog") = method
    ' updateLog
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/logs/:log_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "log"
    Set method("resources") = resources
    Set methods("updateLog") = method
    ' deleteLog
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/logs/:log_id"
    method("verb") = "DELETE"
    Set methods("deleteLog") = method
    ' resetLogs
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/logs"
    method("verb") = "DELETE"
    Set methods("resetLogs") = method
    ' exportLogs
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/logs/export"
    method("verb") = "GET"
    Set methods("exportLogs") = method
    ' listTagsForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/tags"
    method("verb") = "GET"
    Set methods("listTagsForSite") = method
    ' getTagForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/tags/:tag_id"
    method("verb") = "GET"
    Set methods("getTagForSite") = method
    ' renameTagForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/tags/:tag_id"
    method("verb") = "PUT"
    Set methods("renameTagForSite") = method
    ' deleteTagForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/tags/:tag_id"
    method("verb") = "DELETE"
    Set methods("deleteTagForSite") = method
    ' listThemes
    Set method = New Scripting.Dictionary
    method("route") = "/themes"
    method("verb") = "GET"
    Set methods("listThemes") = method
    ' listThemesForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/themes"
    method("verb") = "GET"
    Set methods("listThemesForSite") = method
    ' getTheme
    Set method = New Scripting.Dictionary
    method("route") = "/themes/:theme_id"
    method("verb") = "GET"
    Set methods("getTheme") = method
    ' getThemeForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/themes/:theme_id"
    method("verb") = "GET"
    Set methods("getThemeForSite") = method
    ' applyThemeToSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/themes/:theme_id/apply"
    method("verb") = "POST"
    Set methods("applyThemeToSite") = method
    ' uninstallTheme
    Set method = New Scripting.Dictionary
    method("route") = "/themes/:theme_id"
    method("verb") = "DELETE"
    Set methods("uninstallTheme") = method
    ' exportSiteTheme
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/export_theme"
    method("verb") = "POST"
    Set methods("exportSiteTheme") = method
    ' listTemplates
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates"
    method("verb") = "GET"
    Set methods("listTemplates") = method
    ' getTemplate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id"
    method("verb") = "GET"
    Set methods("getTemplate") = method
    ' createTemplate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "template"
    Set method("resources") = resources
    Set methods("createTemplate") = method
    ' updateTemplate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "template"
    Set method("resources") = resources
    Set methods("updateTemplate") = method
    ' deleteTemplate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id"
    method("verb") = "DELETE"
    Set methods("deleteTemplate") = method
    ' publishTemplate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/publish"
    method("verb") = "POST"
    Set methods("publishTemplate") = method
    ' refreshTemplate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/refresh"
    method("verb") = "POST"
    Set methods("refreshTemplate") = method
    ' refreshTemplatesForSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/refresh_templates"
    method("verb") = "POST"
    Set methods("refreshTemplatesForSite") = method
    ' cloneTemplate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/clone"
    method("verb") = "POST"
    Set methods("cloneTemplate") = method
    ' previewTemplate
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/preview"
    method("verb") = "POST"
    Set methods("previewTemplate") = method
    ' previewTemplateById
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/preview"
    method("verb") = "POST"
    Set methods("previewTemplateById") = method
    ' listTemplatemaps
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/templatemaps"
    method("verb") = "GET"
    Set methods("listTemplatemaps") = method
    ' getTemplatemap
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/templatemaps/:templatemap_id"
    method("verb") = "GET"
    Set methods("getTemplatemap") = method
    ' createTemplatemap
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/templatemaps"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "templatemap"
    Set method("resources") = resources
    Set methods("createTemplatemap") = method
    ' updateTemplatemap
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/templatemaps/:templatemap_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "templatemap"
    Set method("resources") = resources
    Set methods("updateTemplatemap") = method
    ' deleteTemplatemap
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/templates/:template_id/templatemaps/:templatemap_id"
    method("verb") = "DELETE"
    Set methods("deleteTemplatemap") = method
    ' listWidgetsets
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgetsets"
    method("verb") = "GET"
    Set methods("listWidgetsets") = method
    ' getWidgetset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgetsets/:widgetset_id"
    method("verb") = "GET"
    Set methods("getWidgetset") = method
    ' createWidgetset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgetsets"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "widgetset"
    Set method("resources") = resources
    Set methods("createWidgetset") = method
    ' updateWidgetset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgetsets/:widgetset_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "widgetset"
    Set method("resources") = resources
    Set methods("updateWidgetset") = method
    ' deleteWidgetset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgetsets/:widgetset_id"
    method("verb") = "DELETE"
    Set methods("deleteWidgetset") = method
    ' listWidgets
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgets"
    method("verb") = "GET"
    Set methods("listWidgets") = method
    ' listWidgetsForWidgetset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgetsets/:widgetset_id/widgets"
    method("verb") = "GET"
    Set methods("listWidgetsForWidgetset") = method
    ' getWidgets
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgets/:widget_id"
    method("verb") = "GET"
    Set methods("getWidgets") = method
    ' getWidgetForWidgetset
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgetsets/:widgetset_id/widgets/:widget_id"
    method("verb") = "GET"
    Set methods("getWidgetForWidgetset") = method
    ' createWidget
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgets"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "widget"
    Set method("resources") = resources
    Set methods("createWidget") = method
    ' updateWidget
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgets/:widget_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "widget"
    Set method("resources") = resources
    Set methods("updateWidget") = method
    ' deleteWidget
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgets/:widget_id"
    method("verb") = "DELETE"
    Set methods("deleteWidget") = method
    ' refreshWidget
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgets/:widget_id/refresh"
    method("verb") = "POST"
    Set methods("refreshWidget") = method
    ' cloneWidget
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/widgets/:widget_id/clone"
    method("verb") = "POST"
    Set methods("cloneWidget") = method
    ' listUsers
    Set method = New Scripting.Dictionary
    method("route") = "/users"
    method("verb") = "GET"
    Set methods("listUsers") = method
    ' createUser
    Set method = New Scripting.Dictionary
    method("route") = "/users"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "user"
    Set method("resources") = resources
    Set methods("createUser") = method
    ' deleteUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id"
    method("verb") = "DELETE"
    Set methods("deleteUser") = method
    ' unlockUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id/unlock"
    method("verb") = "POST"
    Set methods("unlockUser") = method
    ' recoverPasswordForUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id/recover_password"
    method("verb") = "POST"
    Set methods("recoverPasswordForUser") = method
    ' recoverPassword
    Set method = New Scripting.Dictionary
    method("route") = "/recover_password"
    method("verb") = "POST"
    Set methods("recoverPassword") = method
    ' listPlugins
    Set method = New Scripting.Dictionary
    method("route") = "/plugins"
    method("verb") = "GET"
    Set methods("listPlugins") = method
    ' getPlugin
    Set method = New Scripting.Dictionary
    method("route") = "/plugins/:plugin_id"
    method("verb") = "GET"
    Set methods("getPlugin") = method
    ' enablePlugin
    Set method = New Scripting.Dictionary
    method("route") = "/plugins/:plugin_id/enable"
    method("verb") = "POST"
    Set methods("enablePlugin") = method
    ' disablePlugin
    Set method = New Scripting.Dictionary
    method("route") = "/plugins/:plugin_id/disable"
    method("verb") = "POST"
    Set methods("disablePlugin") = method
    ' enableAllPlugins
    Set method = New Scripting.Dictionary
    method("route") = "/plugins/enable"
    method("verb") = "POST"
    Set methods("enableAllPlugins") = method
    ' disableAllPlugins
    Set method = New Scripting.Dictionary
    method("route") = "/plugins/disable"
    method("verb") = "POST"
    Set methods("disableAllPlugins") = method
    ' backupSite
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/backup"
    method("verb") = "GET"
    Set methods("backupSite") = method
    ' restoreSite
    Set method = New Scripting.Dictionary
    method("route") = "/restore"
    method("verb") = "POST"
    Set methods("restoreSite") = method
    ' listFields
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/fields"
    method("verb") = "GET"
    Set methods("listFields") = method
    ' getField
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/fields/:field_id"
    method("verb") = "GET"
    Set methods("getField") = method
    ' createField
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/fields"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "field"
    Set method("resources") = resources
    Set methods("createField") = method
    ' updateField
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/fields/:field_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "field"
    Set method("resources") = resources
    Set methods("updateField") = method
    ' deleteField
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/fields/:field_id"
    method("verb") = "DELETE"
    Set methods("deleteField") = method
    ' listGroups
    Set method = New Scripting.Dictionary
    method("route") = "/groups"
    method("verb") = "GET"
    Set methods("listGroups") = method
    ' listGroupsForUser
    Set method = New Scripting.Dictionary
    method("route") = "/users/:user_id/groups"
    method("verb") = "GET"
    Set methods("listGroupsForUser") = method
    ' getGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id"
    method("verb") = "GET"
    Set methods("getGroup") = method
    ' createGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "group"
    Set method("resources") = resources
    Set methods("createGroup") = method
    ' updateGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "group"
    Set method("resources") = resources
    Set methods("updateGroup") = method
    ' deleteGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id"
    method("verb") = "DELETE"
    Set methods("deleteGroup") = method
    ' synchronizeGroups
    Set method = New Scripting.Dictionary
    method("route") = "/groups/synchronize"
    method("verb") = "POST"
    Set methods("synchronizeGroups") = method
    ' listPermissionsForGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id/permissions"
    method("verb") = "GET"
    Set methods("listPermissionsForGroup") = method
    ' grantPermissionToGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id/permissions/grant"
    method("verb") = "POST"
    Set methods("grantPermissionToGroup") = method
    ' revokePermissionFromGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id/permissions/revoke"
    method("verb") = "POST"
    Set methods("revokePermissionFromGroup") = method
    ' listMembersForGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id/members"
    method("verb") = "GET"
    Set methods("listMembersForGroup") = method
    ' getMemberForGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id/members/:member_id"
    method("verb") = "GET"
    Set methods("getMemberForGroup") = method
    ' addMemberToGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id/members"
    method("verb") = "POST"
    Set methods("addMemberToGroup") = method
    ' removeMemberFromGroup
    Set method = New Scripting.Dictionary
    method("route") = "/groups/:group_id/members/:member_id"
    method("verb") = "DELETE"
    Set methods("removeMemberFromGroup") = method
    ' bulkAuthorImport
    Set method = New Scripting.Dictionary
    method("route") = "/users/import"
    method("verb") = "POST"
    Set methods("bulkAuthorImport") = method
    ' bulkAuthorExport
    Set method = New Scripting.Dictionary
    method("route") = "/users/export"
    method("verb") = "GET"
    Set methods("bulkAuthorExport") = method
    ' synchronizeUsers
    Set method = New Scripting.Dictionary
    method("route") = "/users/synchronize"
    method("verb") = "POST"
    Set methods("synchronizeUsers") = method
    ' listFormattedTexts
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/formatted_texts"
    method("verb") = "GET"
    Set methods("listFormattedTexts") = method
    ' getFormattedText
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/formatted_texts/:formatted_text_id"
    method("verb") = "GET"
    Set methods("getFormattedText") = method
    ' createFormattedText
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/formatted_texts"
    method("verb") = "POST"
    Set resources = New Collection
    resources.Add "formatted_text"
    Set method("resources") = resources
    Set methods("createFormattedText") = method
    ' updateFormattedText
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/formatted_texts/:formatted_text_id"
    method("verb") = "PUT"
    Set resources = New Collection
    resources.Add "formatted_text"
    Set method("resources") = resources
    Set methods("updateFormattedText") = method
    ' deleteFormattedText
    Set method = New Scripting.Dictionary
    method("route") = "/sites/:site_id/formatted_texts/:formatted_text_id"
    method("verb") = "DELETE"
    Set methods("deleteFormattedText") = method
End Sub

' デストラクタ
Private Sub Class_Terminate()
    Set jlib = Nothing
    Set js = Nothing
    Set fso = Nothing
    Set methods = Nothing
End Sub

' 初期化
Public Sub init(url As String, clientId As String)
    baseUrl = url
    clId = clientId
End Sub

' エンドポイント追加
Public Sub generateEndpointMethod(name As String, method As Scripting.Dictionary)
    methods(name) = method
End Sub

' Data API
Public Function send(ParamArray params()) As Scripting.Dictionary
    Dim method As String, route As String, url As String, verb As String
    Dim sendParams As Scripting.Dictionary, keys() As Variant
    Dim resources As Collection
    Dim re As RegExp, mc As MatchCollection, m As Match, found As String, repl As String
    Dim response As Scripting.Dictionary, items As Collection
    Dim i As Integer, ptr As Integer
    
    ' メソッド名をチェック
    method = params(0)
    If Not methods.Exists(method) Then
        Err.Raise Number:=20999, Description:="メソッド" & method & "は存在しません。"
        Exit Function
    End If
    ' URLを求める
    ptr = 1
    route = methods(method)("route")
    Set re = New RegExp
    re.Global = True
    re.Pattern = "(:[a-z_]+)"
    Set mc = re.Execute(route)
    For i = 0 To mc.Count - 1
        Set m = mc(i)
        found = m.SubMatches(0)
        repl = params(ptr)
        route = Replace(route, found, repl)
        ptr = ptr + 1
    Next
    url = baseUrl & "/v2" & route
    ' resourceをparamsに追加する
    Set sendParams = New Scripting.Dictionary
    If methods(method).Exists("resources") Then
        Set resources = methods(method)("resources")
        For i = 1 To resources.Count
            sendParams(resources(i)) = jlib.toString(params(ptr))
            ptr = ptr + 1
        Next
    End If
    ' その他のパラメータをparamsに追加する
    If ptr = UBound(params) Then
        keys = params(ptr).keys
        For i = 0 To params(ptr).Count - 1
            sendParams(keys(i)) = params(ptr)(keys(i))
        Next
        If method = "authenticate" Then
            sendParams("clientId") = clId
        End If
    End If
    ' Data APIと通信する
    verb = methods(method)("verb")
    Set response = SendData(url, verb, sendParams)
    ' 認証情報の保存／破棄
    If method = "authenticate" Or method = "token" Then
        accessToken = response("accessToken")
    End If
    
    ' itemsをCollectionに変換
    Set items = New Collection
    If response.Exists("items") Then
        For i = 1 To response("items").Count
            items.Add response("items")(i)
        Next
        Set response("items") = items
    End If

    ' 結果を返す
    Set send = response
End Function

' サーバーと通信
Private Function SendData(url As String, verb As String, params As Scripting.Dictionary) As Scripting.Dictionary
    Dim i As Integer, n As String, v As String, boundary As String
    Dim xmlHTTP As MSXML2.XMLHTTP60
    Dim strm As Stream, f_strm As Stream, formData, json, binFile, fName As String
    Dim response As Scripting.Dictionary, items As Collection

    Set xmlHTTP = New MSXML2.XMLHTTP60
    boundary = "----boundary" & Format(Int(Rnd(1) * 100000000), "00000000")
    
    If verb = "GET" Then
        If Not params Is Nothing Then
            For i = 0 To params.Count - 1
                n = params.keys(i)
                v = params(n)
                formData = formData & js.CodeObject.encodeURIComponent(n) & "="
                formData = formData & js.CodeObject.encodeURIComponent(v)
                If i <> params.Count - 1 Then
                    formData = formData & "&"
                End If
            Next
        End If
    Else
        If Not params Is Nothing Then
            ' ストリームを開く
            Set strm = New Stream
            strm.Type = adTypeText
            strm.Charset = "UTF-8"
            strm.Open
            ' パラメータと値をストリームに書き込む
            For i = 0 To params.Count - 1
                n = params.keys(i)
                v = params(n)
                If n = "file" Then
                    Set f_strm = New Stream
                    f_strm.Type = adTypeBinary
                    f_strm.Open
                    f_strm.LoadFromFile v
                    binFile = f_strm.Read()
                    f_strm.Close
                    Set f_strm = Nothing
                    ChangeStreamType strm, adTypeText
                    strm.WriteText "--" & boundary & vbCrLf
                    strm.WriteText "Content-Disposition: form-data; name=""file""; filename=""" & fso.GetFileName(v) & """" & vbCrLf
                    strm.WriteText "Content-Transfer-Encoding: binary" & vbCrLf & vbCrLf
                    ChangeStreamType strm, adTypeBinary
                    strm.Write binFile
                    ChangeStreamType strm, adTypeText
                    strm.WriteText vbCrLf
                Else
                    ChangeStreamType strm, adTypeText
                    strm.WriteText "--" & boundary & vbCrLf
                    strm.WriteText "Content-Disposition: form-data; name=""" & n & """" & vbCrLf & vbCrLf
                    strm.WriteText v & vbCrLf
                End If
            Next
            ChangeStreamType strm, adTypeText
            strm.WriteText "--" & boundary & "--" & vbCrLf
        
            ' ストリームを読み込む
            ChangeStreamType strm, adTypeBinary
            strm.Position = 0
            formData = strm.Read
            strm.Close
            Set strm = Nothing
            
            Open "send.dat" For Binary As #1
            Put #1, , formData
            Close #1
        End If
    End If

    ' データを送信する
    'url = baseUrl & "/" & url
    If verb = "GET" Then
        If formData <> "" Then
            url = url & "?" & formData
        End If
    End If
    xmlHTTP.Open verb, url, False
    If accessToken <> "" Then
        xmlHTTP.setRequestHeader "X-MT-Authorization", "MTAuth accessToken=" & accessToken
    End If
    If verb = "POST" And VarType(formData) = vbByte + vbArray Then
        xmlHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=""" & boundary & """"
        xmlHTTP.setRequestHeader "Content-Length", UBound(formData) + 1
        xmlHTTP.send formData
    Else
        xmlHTTP.send
    End If
    Set response = jlib.parse(xmlHTTP.responseText)
    
    ' 結果を返す
    Set SendData = response
End Function

' itemsをCollectionに変換
Sub ConvertToCollection(ByRef r As Object)
    Dim items As Collection, i As Integer

    Set items = New Collection
    If r("items").Count > 0 Then
        For i = 1 To r("items").Count
            items.Add r("items")(i)
        Next
    End If
    Set r("items") = items
End Sub

' ストリームのタイプを変更
Function ChangeStreamType(st As Stream, t)
    Dim p
    p = st.Position
    st.Position = 0
    st.Type = t
    st.Position = p
    Set ChangeStreamType = st
End Function

Function convDate(d As Date) As String
    convDate = Format(d, "yyyy-mm-ddTHH:MM:SS+09:00")
End Function

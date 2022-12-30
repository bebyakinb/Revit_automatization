"""
Looking for Revit Links then looking at the folder they came from
to see, if there is a higher numbered of revision and if so, then
it will swap to that revision.
"""


import datetime
import fnmatch
import os
import os.path
import subprocess as sp
import time


def get_all_links():
    """Collecting all links, except nested"""
    links = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_RvtLinks)
        # Only Link types, NOT a Nested link
        .Where(lambda link: (RevitLinkType == type(link)) and not link.IsNestedLink)
    )

    links = links.Where(lambda link: Element.Name.GetValue(link).Contains('-RVT-'))

    return links.ToList()


def get_link_name(link):
    return Element.Name.GetValue(link)


def get_link_folder_path(link):
    path = ''
    # Get the External Reference
    er = link.GetExternalFileReference()
    # Make sure it isn't Null
    if not er:
        print(get_link_name(link))
    else:
        # Get the Model Path
        mp = er.GetPath()
        # Convert the Model Path to a human-readable format
        # Also convert relative paths to absolute paths
        model_path = model_pathUtils.Convertmodel_pathToUserVisiblePath(mp)
        path = os.path.dirname(os.path.abspath(model_path))

    return path


def get_revision_num_from_filename(filename):
    link_info = filename.split('-')
    if 'RVT' in link_info:
        # return the item in the list after 'RVT'
        # which should be the revision number.
        revision_num = int(link_info[link_info.IndexOf('RVT') + 1])
    else:
        revision_num = 0

    return revision_num


def get_latest_revision_from_folder(link):
    # Get the path that the Link comes from
    folder_path = get_link_folder_path(link)
    filename = Element.Name.GetValue(link)
    # Only check if the file is from the Revit Links Folder
    if folder_path.Contains('Revit Links'):

        # Get the start of the filename up to the Revision
        # so that we can filter files we are interested in.
        filename_start = filename[:filename.find('-RVT-')]

        # list of Revit files in the folder that match
        # both the start of the filename and end with .rvt
        filenames = fnmatch.filter(os.listdir(folder_path), filename_start + '*.rvt')

        highest = 0
        file = None
        print(folder_path)
        # If more than one file... loop through and find the highest revision
        for filename in filenames:
            num = get_revision_num_from_filename(filename)
            if num > highest:
                highest = num
                file = filename
            else:
                # We have a match
                if num == highest:
                    file = filename
        # Return an updated File Path of a higher Revision
        return folder_path, file
    else:
        # Return an empty string if not from 'Revit Links Folder'.
        return ''


def get_linked_documents():
    documents = {}

    for doc in doc.Application.Documents:
        documents.update({doc.Title: doc})

    return documents


def get_linked_document_doc(documents, link_name):
    # remove the ".rvt"
    link_name = link_name[:link_name.find('.rvt')]

    # Check that the document is in the open documents (ie a Link)
    if documents.ContainsKey(link_name):
        doc_link = documents[link_name]
    else:
        doc_link = None
    return doc_link


def get_closed_workset_ids(doc_link):
    # Get the Worksets
    worksets = (
        FilteredWorksetCollector(doc_link)
        .Where(lambda ws: ws.Kind == WorksetKind.UserWorkset)
    )

    closed_worksets = []

    for workset in worksets:
        if not workset.IsOpen:
            closed_worksets.Add(workset.Id)

    # Return the list of Ids
    return closed_worksets


def log_list_closed_worksets(doc_link):
    wst = doc_link.GetWorksetTable()

    # Get the List of Closed Worksets
    closed_worksets = get_closed_workset_ids(doc_link)
    wc = WorksetConfiguration()
    wc.Close(closed_worksets)

    log = '\nClosed Worksets as follows:'
    # Print the list of Closed Worksets to the Log
    for workset_name in closed_worksets:
        log += '\n  [-] ' + wst.GetWorkset(workset_name).Name
    return log


def create_log_file(path, text):
    # Establish the Log Filename
    model_name = doc.Title[:doc.Title.LastIndexOf('_')]
    timestamp = datetime.datetime.now().strftime("(%Y-%m-%d - %H-%M-%S)")
    file_path = path + '\\' + model_name + '_ReloadLinks_' + timestamp + '.log'

    # Open, Write to the File & Close
    fp = open(file_path, 'w')
    fp.write(text)
    fp.close()

    # Then Open in Notepad.exe
    logfile = path
    sp.Popen(['notepad.exe', logfile])


def relink_from_revision(link, bad_models):
    log = ''
    model_path = model_pathUtils.ConvertUserVisiblePathTomodel_path(new_file_path)
    try:
        link.LoadFrom(model_path, WorksetConfiguration())
    except Exception as e:
        bad_models.append((get_link_name(link), e.message))
        log += "ERROR LOADING MODEL"
        log += log + e.message
    return log


def update_links(not_workshared_links, not_found_doc_links, bad_models):
    # Get all the current open documents
    documents = get_linked_documents()

    # Get all the Links (LinkTypes Only)
    links = get_all_links()

    # Create a Log Element for each link
    log_separator = '+' + '-' * 80 + '+'
    log = '\n' + log_separator

    # For Each of the Links:
    for link in links:
        link_path = get_link_folder_path(link)
        link_name = get_link_name(link)

        if 'Revit Links' in link_path:
            log += '\nLink Name = ' + link_name
            log += '\nLink Path = ' + link_path

            # Check to see if this link is loaded, so we can re-instate this status
            is_loaded = link.IsLoaded(documents, link.Id)
            # If the Link was unloaded, reload it while we check for updates.
            # It will be unloaded at the end.
            if not is_loaded:
                link.Reload()

            # Get the Current Revision of this link
            current_revision_link = get_revision_num_from_filename(link_name)

            # Get the Highest Revision Available
            new_link_path, new_link_name = get_latest_revision_from_folder(link)
            new_file_path = new_link_path + '\\' + new_link_name
            # Get the Potential New Revision Number
            new_revision = get_revision_num_from_filename(new_link_name)

            if current_revision_link.Equals(new_revision):
                # No Relinking is required.
                log += '\nFile: ' + link_name + ' is already up to date at Rev ' + current_revision_link.ToString()
            else:
                # We need to relink from the newer revision.
                log += '\nNeed to update ' + current_revision_link.ToString() + ' to Rev ' + new_revision.ToString()

                doc_link = get_linked_document_doc(documents, link_name)

                # make Sure we got a valid document back.
                if doc_link:
                    if doc_link.IsWorkshared:
                        # If the link is Workshared we need to make sure we match any closed Worksets
                        log += log_list_closed_worksets(doc_link)

                        # Relink From the New Revision with Workset Control
                        model_path = model_pathUtils.ConvertUserVisiblePathTomodel_path(new_file_path)
                        link.LoadFrom(model_path, None)
                    else:
                        not_workshared_links.Add(get_link_name(link))
                        log += '\nThis file is NOT Workshared...!'
                        # Relink From the New Revision (No Workset Control)
                        log += relink_from_revision(link, bad_models)

                else:
                    log += '\ndoc_link Not Found'
                    not_found_doc_links.Add(get_link_name(link))

            log += '\n' + log_separator

        else:
            pass

    return log


def main():
    not_workshared_links = []
    not_found_doc_links = []
    bad_models = []

    log = update_links(not_workshared_links, not_found_doc_links, bad_models)
    # List any non-Workshared links
    if not_workshared_links.Count > 0:
        log += '\nThe Following Links are not workshared...!!\n'
        for nwl in not_workshared_links:
            log += '\n' + nwl

    # List any files that weren't found and need manual attention
    if not_found_doc_links.Count > 0:
        log += '\nThe Following Links were not found.'
        for not_found_doc_link in not_found_doc_links:
            log += '\n' + not_found_doc_link

    if bad_models.Count > 0:
        log += '\nThe Following Links could not be loaded.'
        for m in bad_models:
            log += '\n' + m[0]
            log += '\n' + m[1]

    # Print the Link Log to the Log File
    create_log_file(log)


if __name__ == "__main__":
    main()

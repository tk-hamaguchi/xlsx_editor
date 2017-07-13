require 'fileutils'
require 'rexml/document'

require 'zip'

require 'xlsx_editor/version'
require 'xlsx_editor/commons'
require 'xlsx_editor/fetcher'
require 'xlsx_editor/updater'

module XlsxEditor
  include Commons, Fetcher, Updater

  module_function *%i[fetch update sheet_names_for shared_items_for]
end

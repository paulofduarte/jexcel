/**
 * (c) 2013 Jexcel Plugin v1.0.0 > Bossanova UI
 * http://www.github.com/paulhodel/jexcel
 *
 * @author: Paul Hodel <paul.hodel@gmail.com>
 * @description: Create light embedded spreadsheets on your webpages
 * 
 * ROADMAP:
 * spare rows and columns
 * Multiple tabs
 * Merged cells
 * Reorder methods
 * Drag and drop rows and columns
 * Custom renderer
 * big data (partial table loading)
 * indexing
 * ctrl+z
 * Context menu
 * Initial loading using setData methods
 * Custom render
 */

(function( $ ){

var methods = {

    /**
     * Innitialization, configuration and loading
     * 
     * @param {Object} options configuration
     * @return void
     */
    init : function( options ) {
        // Loading default configuration
        var defaults = {
            colHeaders:[],
            colWidths:[],
            colAlignments:[],
            columns:[],
        };

        // Configuration holder
        var options =  $.extend(defaults, options);

        // Id
        var id = $(this).prop('id');

        // Main object
        var main = $(this);

        // Create
        prepareTable = function () {
            // Defaults to avoid erros
            if (! options.colHeaders.length) {
                for (i = 0; i < options.data[0].length; i++) {
                    options.colHeaders[i] = '';
                }
            }
            if (! options.columns.length) {
                for (i = 0; i < options.data[0].length; i++) {
                    options.columns[i] = { type:'text' };
                    options.colAlignments[i] = 'center';
                }
            }

            // Register options
            if (! $.fn.jexcel.defaults) {
                $.fn.jexcel.defaults = new Array();
            }
            $.fn.jexcel.defaults[id] = options;

            // Loading initial data from remote sources
            var results = [];

            // Number of columns
            size = options.colHeaders.length;
            if (options.data[0]) {
                if (options.data[0].length > size) {
                    size = options.data[0].length;
                }
            }

            // Preparations
            for (i = 0; i < size; i++) {
                // Avoid erros in case the configuration is not complete
                if (! options.colHeaders[i]) {
                    options.colHeaders[i] = '';
                }
                if (! options.columns[i]) {
                    options.columns[i] = { type:'text' };
                }
                if (! options.colAlignments[i]) {
                    options.colAlignments[i] = 'center';
                }
                if (! options.columns[i].source) {
                    $.fn.jexcel.defaults[id].columns[i].source = [];
                }

                // Pre-load initial source for json autocomplete
                if (options.columns[i].type == 'autocomplete' || options.columns[i].type == 'dropdown') {
                    // if remote content
                    if (options.columns[i].url) {
                        results.push($.ajax({
                            url: options.columns[i].url,
                            index: i,
                            dataType:'json',
                            success: function (result) {
                                // Create the dynamic sources
                                $.fn.jexcel.defaults[id].columns[this.index].source = result;
                                // Populate the combo variable
                                $.fn.jexcel.defaults[id].columns[this.index].combo = $(main).jexcel('createCombo', result);
                            }
                        }));
                    } else if (options.columns[i].source) {
                        // Create the dropdown combo based on the source
                        $.fn.jexcel.defaults[id].columns[i].combo = $(main).jexcel('createCombo', options.columns[i].source);
                    }
                } else if (options.columns[i].type == 'calendar') {
                    // Avoid erros in case configuration is not complete
                    if (! $.fn.jexcel.defaults[id].columns[i].options) {
                        $.fn.jexcel.defaults[id].columns[i].options = [];
                    }
                    // Default format for date columns
                    if (! $.fn.jexcel.defaults[id].columns[i].options.format) {
                        $.fn.jexcel.defaults[id].columns[i].options.format = 'DD/MM/YYYY';
                    }
                } else {
                    // Check if a mask is requested for a specific column
                    if ($.fn.jexcel.defaults[id].columns[i].mask) {
                        // Check if the third party plugin is loaded in the page
                        if ($.fn.masked) {
                            if (! $.fn.jexcel.defaults[id].columns[i].options) {
                                $.fn.jexcel.defaults[id].columns[i].options = [];
                            }
                        }
                    }
                }
            }

            // In case there are external json to be loaded before create the table
            if (results.length > 0) {
                // Waiting all external data is loaded
                $.when.apply(this, results).done(function() {
                    // Create the table
                    $(main).jexcel('createTable');
                });
            } else {
                // No external data to be loaded, just created the table
                $(main).jexcel('createTable');
            }
        }

        // Load the table data based on an CSV file
        if (options.csv) {
            if (! $.csv) {
                // Required lib not present
                console.error('Jexcel error: jquery-csv library not loaded');
            } else {
                // Comma as default
                options.delimiter = options.delimiter || ',';

                // Load CSV file
                $.ajax({
                    url: options.csv,
                    success: function (result) {
                        var i = 0;
                        // Converted data
                        options.data = $.csv.toArrays(result);
                        // Prepare table
                        prepareTable();
                    }
                });
            }
        } else {
            // Prepare table
            prepareTable();
        }
    },

    /**
     * Create the table
     * 
     * @return void
     */
    createTable : function() {
        // Id
        var id = $(this).prop('id');

        // Data
        if (! $.fn.jexcel.defaults[id].data) {
            $.fn.jexcel.defaults[id].data = [];
        }

        // Length
        if (! $.fn.jexcel.defaults[id].data.length) {
            $.fn.jexcel.defaults[id].data = [[]];
        }

        // Var options
        var options = $.fn.jexcel.defaults[id];

        // Create main table object
        var table = document.createElement('table');
        $(table).prop('class', 'jexcel bossanova-ui');
        $(table).prop('cellpadding', '0');
        $(table).prop('cellspacing', '0');

        // Unselectable properties
        $(table).prop('unselectable', 'yes');
        $(table).prop('onselectstart', 'return false');
        $(table).prop('draggable', 'false');

        // Create header and body tags
        var thead = document.createElement('thead');
        var tbody = document.createElement('tbody');

        // Header
        $(thead).prop('class', 'label');

        // Create headers
        var tr = '<td width="30" class="label"></td>';

        for (i = 0; i < options.colHeaders.length; i++) {
            // Default column should be text
            if (! options.columns[i]) {
                options.columns[i] = { type: 'text' };
            }

            // If no header title is defined described column with a letter
            if (! options.colHeaders[i]) {
                options.colHeaders[i] = '';
                if (i > 701) {
                    options.colHeaders[i] += String.fromCharCode(64 + parseInt(i / 676));
                    options.colHeaders[i] += String.fromCharCode(64 + parseInt((i % 676) / 26));
                } else if (i > 25) {
                    options.colHeaders[i] += String.fromCharCode(64 + parseInt(i / 26));
                }
                options.colHeaders[i] += String.fromCharCode(65 + (i % 26));
            }

            // Default header cell properties
            width = options.colWidths[i] || 50;
            align = options.colAlignments[i] || 'left';

            // Column type hidden
            if (options.columns[i].type == 'hidden') {
                // TODO: when it is first check the whole selection not include
                tr += '<td id="col-' + i + '" style="display:none;">' + options.colHeaders[i] + '</td>';
            } else {
                // Other column types
                tr += '<td id="col-' + i + '" width="' + width + '" align="' + align +'" title="' + options.colHeaders[i] + '">' + options.colHeaders[i] + '</td>';
            }
        }

        // Populate header
        $(thead).html('<tr>' + tr + '</tr>'); 

        // TODO: filter row
        //<tr><td></td><td><input type="text"></td></tr>

        // Append content
        $(table).append(thead);
        $(table).append(tbody);

        // Prevent dragging
        $(table).on('dragstart', function () {
            return false;
        });

        // Main object
        $(this).html(table);

        // Add the corner square and textarea one time onlly
        if (! $('.jexcel_corner').length) {
            // Corner one for all sheets in a page
            var corner = document.createElement('div');
            $(corner).prop('class', 'jexcel_corner');
            $(corner).prop('id', 'corner');

            // Hidden textarea copy and paste helper
            var textarea = document.createElement('textarea');
            $(textarea).prop('class', 'jexcel_textarea');
            $(textarea).prop('id', 'textarea');

            // Powered by
            var ads = document.createElement('div');
            $(ads).css('display', 'none');
            $(ads).html('<a href="http://github.com/paulhodel/jexcel">jExcel Spreadsheet</a>');

            // Append elements
            $('body').append(corner);
            $('body').append(textarea);
            $('body').append(ads);

            // Prevent dragging on the corner object
            $(corner).on('dragstart', function () {
                return false;
            });

            // Corner persistence
            $.fn.jexcel.selectedCorner = false;
            $.fn.jexcel.selectedHeader = null;

            // Global mouse click down controles
            $(document).on('mousedown', function (e) {
                // Click on corner icon
                if (e.target.id == 'corner') {
                    $.fn.jexcel.selectedCorner = true;
                } else {
                    // Check if the click was in an jexcel element
                    var table = $(e.target).parent().parent().parent();

                    // Table found
                    if ($(table).is('.jexcel')) {

                        // Get id
                        var current = $(table).parent().prop('id');

                        // Remove selection from any other jexcel if applicable
                        if ($.fn.jexcel.current) {
                            if ($.fn.jexcel.current != current) {
                                $('#' + $.fn.jexcel.current).find('td').removeClass('selected highlight highlight-top highlight-left highlight-right highlight-bottom');
                            }
                        }

                        // Mark as current
                        $.fn.jexcel.current = current;

                        // Header found
                        if ($(e.target).parent().parent().is('thead')) {
                            var o = $(e.target).prop('id');
                            if (o) {
                                o = o.split('-');

                                if ($.fn.jexcel.selectedHeader && (e.shiftKey || e.ctrlKey)) {
                                    var d = $($.fn.jexcel.selectedHeader).prop('id').split('-');
                                } else {
                                    // Update selection single column
                                    var d = $(e.target).prop('id').split('-');
                                    // Keep track of which header was selected first
                                    $.fn.jexcel.selectedHeader = $(e.target);
                                }

                                 // Get cell objects 
                                var o1 = $('#' + $.fn.jexcel.current).find('#' + o[1] + '-0');
                                var o2 = $('#' + $.fn.jexcel.current).find('#' + d[1] + '-' + parseInt($.fn.jexcel.defaults[$.fn.jexcel.current].data.length - 1));

                                // Update selection
                                $('#' + $.fn.jexcel.current).jexcel('updateSelection', o1, o2);
                            }
                        } else {
                            $.fn.jexcel.selectedHeader = false;
                        }

                        // Body found
                        if ($(e.target).parent().parent().is('tbody')) {
                            // Update row label selection
                            if ($(e.target).is('.label')) {
                                var o = $(e.target).prop('id').split('-');

                                if ($.fn.jexcel.selectedRow && (e.shiftKey || e.ctrlKey)) {
                                    // Updade selection multi columns
                                    var d = $($.fn.jexcel.selectedRow).prop('id').split('-');
                                } else {
                                    // Update selection single column
                                    var d = $(e.target).prop('id').split('-');
                                    // Keep track of which header was selected first
                                    $.fn.jexcel.selectedRow = $(e.target);
                                }

                                // Get cell objects 
                                var o1 = $('#' + $.fn.jexcel.current).find('#0-' + o[1]);
                                var o2 = $('#' + $.fn.jexcel.current).find('#' + parseInt($.fn.jexcel.defaults[$.fn.jexcel.current].columns.length - 1) + '-' + d[1]);

                                $('#' + $.fn.jexcel.current).jexcel('updateSelection', o1, o2);
                            } else {
                                // Update cell selection
                                if (! $($.fn.jexcel.selectedCell).hasClass('edition')) {
                                    if (! $.fn.jexcel.selectedCell || ! e.shiftKey) {
                                        $.fn.jexcel.selectedCell = $(e.target);
                                    }
                                    $('#' + $.fn.jexcel.current).jexcel('updateSelection', $.fn.jexcel.selectedCell, $(e.target));
                                } else {
                                    if ($(e.target) != $.fn.jexcel.selectedCell) {
                                        $.fn.jexcel.selectedCell = $(e.target);
                                        $('#' + $.fn.jexcel.current).jexcel('updateSelection', $.fn.jexcel.selectedCell, $(e.target));
                                    }
                                }

                                // No full row selected
                                $.fn.jexcel.selectedRow = null;
                            }
                        }
                    } else {
                        // Remove selection from any other jexcel if applicable
                        if ($.fn.jexcel.current) {
                            $('#' + $.fn.jexcel.current).find('td').removeClass('selected highlight highlight-top highlight-left highlight-right highlight-bottom');
                        }

                        // Hide corner
                        $(corner).css('top', '-200px');
                        $(corner).css('left', '-200px');

                        // Reset controls
                        $.fn.jexcel.current = null;
                        $.fn.jexcel.selectedCell = null;
                        $.fn.jexcel.selectedRow = null;
                        $.fn.jexcel.selectedHeader = null;
                    }
                }
            });

            // Global mouse click up controles 
            $(document).mouseup(function (o) {
                // Cancel any corner selection
                $.fn.jexcel.selectedCorner = false;

                // Data to be copied
                var selection = $('#' + $.fn.jexcel.current).find('tbody td.selection');

                if ($(selection).length > 0) {
                    // First and last cells
                    var o = $(selection[0]).prop('id').split('-');
                    var d = $(selection[selection.length - 1]).prop('id').split('-');

                    // Copy data
                    $('#' + $.fn.jexcel.current).jexcel('copyData', o, d);

                    // Remove selection
                    $(selection).removeClass('selection selection-left selection-right selection-top selection-bottom');
                }
            });

            // Double click
            $(document).on('dblclick', function (e) {
                // Jexcel is selected
                if ($.fn.jexcel.current) {
                    // Corner action
                    if (e.target.id == 'corner') {
                        var selection = $('#' + $.fn.jexcel.current).find('tbody td.highlight');
                        // Any selected cells
                        if (typeof(selection) == 'object') {
                            // Get selected cells
                            var o = $(selection[0]).prop('id').split('-');
                            var d = $(selection[selection.length - 1]).prop('id').split('-');
                            // Double click copy
                            o[1] = parseInt(d[1]) + 1;
                            d[1] = parseInt($.fn.jexcel.defaults[$.fn.jexcel.current].data.length);
                            // Do copy
                            $('#' + $.fn.jexcel.current).jexcel('copyData', o, d);
                        }
                    }

                    // Open editor action
                    if ($(e.target).is('.highlight')) {
                        $('#' + $.fn.jexcel.current).jexcel('openEditor', $(e.target));
                    }
                }
            });

            $(document).on('mouseover', function (e) {
                // Get jexcel table
                var table = $(e.target).closest('.jexcel');

                // If the user is in the current table
                if ($.fn.jexcel.current == $(table).parent().prop('id')) {
                    // Header found
                    if ($(e.target).parent().parent().is('thead')) {
                        if ($.fn.jexcel.selectedHeader) {
                            // Updade selection
                            if (e.buttons) {
                                var o = $($.fn.jexcel.selectedHeader).prop('id');
                                var d = $(e.target).prop('id');
                                if (o && d) {
                                    o = o.split('-');
                                    d = d.split('-');
                                    // Get cell objects 
                                    var o1 = $('#' + $.fn.jexcel.current).find('#' + o[1] + '-0');
                                    var o2 = $('#' + $.fn.jexcel.current).find('#' + d[1] + '-' + parseInt($.fn.jexcel.defaults[$.fn.jexcel.current].data.length - 1));
                                    // Update selection
                                    $('#' + $.fn.jexcel.current).jexcel('updateSelection', o1, o2);
                                }
                            }
                        }
                    }

                    // Body found
                    if ($(e.target).parent().parent().is('tbody')) {
                        // Update row label selection
                        if ($(e.target).is('.label')) {
                            if ($.fn.jexcel.selectedRow) {
                                // Updade selection
                                if (e.buttons) {
                                    var o = $($.fn.jexcel.selectedRow).prop('id');
                                    var d = $(e.target).prop('id');
                                    if (o && d) {
                                        o = o.split('-');
                                        d = d.split('-');
                                        // Get cell objects 
                                        var o1 = $('#' + $.fn.jexcel.current).find('#0-' + o[1]);
                                        var o2 = $('#' + $.fn.jexcel.current).find('#' + parseInt($.fn.jexcel.defaults[$.fn.jexcel.current].columns.length - 1) + '-'  + d[1]);
                                        // Update selection
                                        $('#' + $.fn.jexcel.current).jexcel('updateSelection', o1, o2);
                                    }
                                }
                            }
                        } else {
                            if ($.fn.jexcel.selectedCell) {
                                if (! $($.fn.jexcel.selectedCell).hasClass('edition')) {
                                    if ($.fn.jexcel.selectedCorner == true) {
                                        // Copy option
                                        $('#' + $.fn.jexcel.current).jexcel('updateCornerSelection', $(e.target));
                                    } else {
                                        // Updade selection
                                        if (e.buttons) {
                                            $('#' + $.fn.jexcel.current).jexcel('updateSelection', $.fn.jexcel.selectedCell, $(e.target));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });

            // Copy data from the table in excel format
            $(document).on('copy', function(e) {
                if ($.fn.jexcel.current) {
                    // Copy data
                    $('#' + $.fn.jexcel.current).jexcel('copy', true);
                }
            });

            // Cut data from the table in excel format
            $(document).on('cut', function() {
                if ($.fn.jexcel.current) {
                    // Copy data
                    $('#' + $.fn.jexcel.current).jexcel('copy', true);
                    // Remove current data 
                    $('#' + $.fn.jexcel.current).jexcel('setValue', $('#' + $.fn.jexcel.current).find('.highlight'), '');
                }
            });

            // Paste data from excel format to the table
            $(document).on('paste', function(e) {
                if ($.fn.jexcel.current) {
                    $('#' + $.fn.jexcel.current).jexcel('paste', $.fn.jexcel.selectedCell, e.originalEvent.clipboardData.getData('text'));
                }
            });

            // Keyboard controls
            var keyBoardCell = null;

            $(document).keydown(function(e) {
                if ($.fn.jexcel.current) {
                    var cell = null;

                    if (e.which == 37) {
                        // Left arrow
                       if (! $($.fn.jexcel.selectedCell).hasClass('edition')) {
                          cell = $($.fn.jexcel.selectedCell).prev();
                       }
                    } else if (e.which == 39) {
                        // Right arrow
                       if (! $($.fn.jexcel.selectedCell).hasClass('edition')) {
                          cell = $($.fn.jexcel.selectedCell).next();
                       }
                    } else if (e.which == 38) {
                        // Top arrow
                       if (! $($.fn.jexcel.selectedCell).hasClass('edition')) {
                          i = $($.fn.jexcel.selectedCell).prop('id').split('-');
                          cell = $($.fn.jexcel.selectedCell).parent().prev().find('#' + i[0] + '-' + (i[1] - 1));
                       }
                    } else if (e.which == 40) {
                        // Bottom arrow
                       if (! $($.fn.jexcel.selectedCell).hasClass('edition')) {
                          i = $($.fn.jexcel.selectedCell).prop('id').split('-');
                          cell = $($.fn.jexcel.selectedCell).parent().next().find('#' + i[0] + '-' + (parseInt(i[1]) + 1));
                       }
                    } else if (e.which == 27) {
                        // Escape
                        if ($($.fn.jexcel.selectedCell).hasClass('edition')) {
                            // Exit without saving
                            $('#' + $.fn.jexcel.current).jexcel('closeEditor', $($.fn.jexcel.selectedCell), false);
                        }
                    } else if (e.which == 13) {
                        // Enter key - Get the id of the selected cell
                        i = $($.fn.jexcel.selectedCell).prop('id').split('-');
                        // Edition in progress
                        if ($($.fn.jexcel.selectedCell).hasClass('edition')) {
                            // Exit saving data
                            if ($.fn.jexcel.defaults[$.fn.jexcel.current].columns[i[0]].type == 'calendar') {
                                $('#' + $.fn.jexcel.current).find('editor').jcalendar('close', 1)
                            } else {
                                $('#' + $.fn.jexcel.current).jexcel('closeEditor', $($.fn.jexcel.selectedCell), true);
                            }
                        } else {
                            // If not edition check if the selected cell is in the last row
                            if (i[1] == $.fn.jexcel.defaults[$.fn.jexcel.current].data.length - 1) {
                                // New record in case selectedCell in the last row
                                $('#' + $.fn.jexcel.current).jexcel('insertRow');
                            }
                        }
                    } else if (e.which == 9) {
                        // Tab key - Get the id of the selected cell
                        i = $($.fn.jexcel.selectedCell).prop('id').split('-');
                        if (i[0] == $.fn.jexcel.defaults[$.fn.jexcel.current].data[0].length - 1) {
                            // New record in case selectedCell in the last column
                            $('#' + $.fn.jexcel.current).jexcel('insertColumn');
                        }
                    } else if (e.which == 46) {
                        // Delete (erase cell in case no edition is running)
                        if (! $($.fn.jexcel.selectedCell).hasClass('edition')) {
                            $('#' + $.fn.jexcel.current).jexcel('setValue', $('#' + $.fn.jexcel.current).find('.highlight'), '');
                        }
                    } else {
                        if (! e.shiftKey && ! e.ctrlKey) {
                            if ($.fn.jexcel.selectedCell) {
                                // Get cell id
                                i = $($.fn.jexcel.selectedCell).prop('id').split('-');
                                // If is not readonly
                                if ($.fn.jexcel.defaults[$.fn.jexcel.current].columns[i[0]].type != 'readonly') {
                                    // Start edition in case a valid character. 
                                    if (! $($.fn.jexcel.selectedCell).hasClass('edition')) {
                                        // TODO: check the sample characters able to start a edition
                                        if (/[a-zA-Z0-9]/.test(String.fromCharCode(e.keyCode))) {
                                            $('#' + $.fn.jexcel.current).jexcel('openEditor', $($.fn.jexcel.selectedCell), true);
                                        }
                                    }
                                }
                            }
                        } else {
                            if (e.which == 65) {
                                // Ctrl + A
                                t = $(this).find('.jexcel tbody td').not('.label');
                                o = $(t).first();
                                t = $(t).last();
                                console.log(o);
                                console.log(t);
                                $('#' + $.fn.jexcel.current).jexcel('updateSelection', o, t);
                                // Prevent page selection
                                e.preventDefault();
                            } else if (e.which == 83) {
                                // Ctrl + S
                                $('#' + $.fn.jexcel.current).jexcel('download');
                                // Prevent page selection
                                e.preventDefault();
                            }
                        }
                    }

                    // Arrows control
                    if (cell) {
                        // Control selected cell
                        if ($(cell).length > 0 && $(cell).prop('id').substr(0,3) != 'row') {
                            // In case of a multiple cell selection
                            if (e.shiftKey || e.ctrlKey) {
                                // Keep first selected cell
                                if (! keyBoardCell) {
                                    keyBoardCell = $.fn.jexcel.selectedCell;
                                }

                                // Origin cell
                                o = keyBoardCell;
                            } else {
                                // Single selection reset history
                                keyBoardCell = null;

                                // Origin cell
                                o = cell;
                            }

                            // Target cell
                            t = cell;

                            // Current cell
                            $.fn.jexcel.selectedCell = cell;

                            // Focus
                            $(cell).focus();

                            // Update selection
                            $('#' + $.fn.jexcel.current).jexcel('updateSelection', o, t);
                        }
                    }
                }
            });
        }

        // Load data
        $(this).jexcel('setData');
    },

    /**
     * Set data
     * 
     * @param array data In case no data is sent, default is reloaded
     * @return void
     */
    setData : function(data) {
        // Id
        var id = $(this).prop('id');

        // Update data
        if (data) {
            if (typeof(data) == 'string') {
                data = JSON.parse(data);
            }

            $.fn.jexcel.defaults[id].data = data;
        }

        // Options
        var options = $.fn.jexcel.defaults[id];

        // Dynamic columns
        $.fn.jexcel.defaults[id].dynamicColumns = [];

        // Data container
        var tbody = $(this).find('tbody');

        for (j = 0; j < options.data.length; j++) {
            // New line of data to be append in the table
            tr = document.createElement('tr');
            // Index column
            $(tr).append('<td id="row-' + j + '" class="label">' + parseInt(j + 1) + '</td>'); 
            // Data columns
            for (i = 0; i < options.colHeaders.length; i++) {
                // New column of data to be append in the line
                td = document.createElement('td');
                // Line properties
                align = options.colAlignments[i] || 'left';
                width = options.colWidths[i] || 50;
                $(td).prop('width', width);
                $(td).prop('align', align);
                $(td).prop('id', i + '-' +j);

                // Readonly
                if (options.columns[i].readOnly == true) {
                    $(td).prop('readonly', 'readonly');
                }
                // Hidden column
                if (options.columns[i].type == 'hidden') {
                    $(td).css('display', 'none');
                }
                // Column value
                val = '' + options.data[j][i];
                $(this).jexcel('setValue', $(td), val, true);

                // Add column to the row
                $(tr).append(td);
            }

            // Add row to the table body
            $(tbody).append(tr);
        }

        // Dynamic updates
        if ($.fn.jexcel.defaults[id].dynamicColumns.length > 0) {
            $(this).jexcel('formula');
        }

        // Table is ready
        if (typeof($.fn.jexcel.defaults[id].onload) == 'function') {
            $.fn.jexcel.defaults[id].onload($(this));
        }
    },

    /**
     * Update table settings helper. Update cells after loading
     * 
     * @param methods
     * @return void
     */
    updateSettings : function(options) {
        // Id
        var id = $(this).prop('id');

        // Keep options
        if (! options) {
            if ($.fn.jexcel.defaults[id].updateSettingsOptions) {
                options = $.fn.jexcel.defaults[id].updateSettingsOptions;
            }
        }

        // Go through all cells
        if (typeof(options) == 'object') {
            $.fn.jexcel.defaults[id].updateSettingsOptions = options;
            var cells = $(this).find('.jexcel tbody td').not('.label');
            if (typeof(options.cells) == 'function') {
                $.each(cells, function (k, v) {
                    id = $(v).prop('id').split('-');
                    options.cells($(v), id[0], id[1]);
                });
            }
        }
    },

    /**
     * Open the editor
     * 
     * @param object cell
     * @return void
     */
    openEditor : function(cell, empty) {
        // Id
        var id = $(this).prop('id');

        // Main
        var main = $(this);

        // Options
        var options = $.fn.jexcel.defaults[id];

        // Get cell position
        var position = $(cell).prop('id').split('-');

        // Readonly
        if ($(cell).hasClass('readonly') == true) {
            // Do nothing
        } else {
            // Holder
            $.fn.jexcel.edition = $(cell).html();

            // If there is a custom editor for it
            if (options.columns[position[0]].editor) {
                // Keep the current value
                $(cell).addClass('edition');

                // Custom editors
                options.columns[position[0]].editor.openEditor(cell);
            } else {
                // Native functions
                if (options.columns[position[0]].type == 'checkbox' || options.columns[position[0]].type == 'hidden') {
                    // Do nothing for checkboxes or hidden columns
                } else if (options.columns[position[0]].type == 'dropdown') {
                    // Keep the current value
                    $(cell).addClass('edition');

                    // Create dropdown
                    var source = options.columns[position[0]].source;

                    var html = '<select>';
                    for (i = 0; i < source.length; i++) {
                        if (typeof(source[i]) == 'object') {
                            k = source[i].id;
                            v = source[i].name;
                        } else {
                            k = source[i];
                            v = source[i];
                        }
                        html += '<option value="' + k + '">' + v + '</option>';
                    }
                    html += '</select>';

                    // Get current value
                    var value = $(cell).find('input').val();

                    // Open editor
                    $(cell).html(html);

                    // Editor configuration
                    var editor = $(cell).find('select');
                    $(editor).change(function () {
                        $(main).jexcel('closeEditor', $(this).parent(), true);
                    });
                    $(editor).blur(function () {
                        $(main).jexcel('closeEditor', $(this).parent(), true);
                    });
                    
                    $(editor).focus();
                    if (value) {
                        $(editor).val(value);
                    }
                } else if (options.columns[position[0]].type == 'calendar') {
                    $(cell).addClass('edition');

                    // Get content
                    var value = $(cell).find('input').val();

                    // Basic editor
                    var editor = document.createElement('input');
                    $(editor).prop('class', 'editor');
                    $(editor).css('width', $(cell).width());
                    $(editor).val($(cell).text());
                    $(cell).html(editor);
                    $(cell).find('');
                    $(cell).focus();

                    options.columns[position[0]].options.onclose = function () {
                        $(main).jexcel('closeEditor', $(cell), true);
                    }

                    // Current value
                    $(editor).jcalendar(options.columns[position[0]].options);
                    $(editor).jcalendar('open', value);
                } else if (options.columns[position[0]].type == 'autocomplete') {
                    // Keep the current value
                    $(cell).addClass('edition');

                    // Get content
                    var html = $(cell).text();
                    var value = $(cell).find('input').val();

                    // Basic editor
                    var editor = document.createElement('input');
                    $(editor).prop('class', 'editor');
                    $(editor).css('width', $(cell).width());

                    // Results
                    var result = document.createElement('div');
                    $(result).prop('class', 'results');
                    if (html) {
                       $(result).html('<li id="' + value + '">' + html + '</li>');
                    } else {
                       $(result).css('display', 'none');
                    }

                    // Search
                    var timeout = null;
                    $(editor).on('keyup', function () {
                        // String
                        var str = $(this).val();

                        // Timeout
                        if (timeout) {
                            clearTimeout(timeout)
                        }

                        // Delay search
                        timeout = setTimeout(function () { 
                            // Object
                            $(result).html('');
                            // List result
                            showResult = function(data, str) {
                                // Create options
                                $.each(data, function(k, v) {
                                    if (typeof(v) == 'object') {
                                        name = v.name;
                                        id = v.id;
                                    } else {
                                        name = v;
                                        id = v;
                                    }

                                    if (name.toLowerCase().indexOf(str.toLowerCase()) != -1) {
                                        li = document.createElement('li');
                                        $(li).prop('id', id)
                                        $(li).html(name);
                                        $(li).mousedown(function (e) {
                                            // TODO: avoid other selection in this handler.
                                            $(cell).html(this);
                                            $(main).jexcel('closeEditor', $(cell), true);
                                        });
                                        $(result).append(li);
                                    }
                                });

                                if (! $(result).html()) {
                                    $(result).html('<div style="padding:6px;">No result found</div>');
                                }
                                $(result).css('display', '');
                            }

                            // Search
                            if (options.columns[position[0]].url) {
                                $.getJSON (options.columns[position[0]].url + '?q=' + str + '&r=' + $(main).jexcel('getRowData', position[1]), function (data) {
                                    showResult(data, str);
                                });
                            } else if (options.columns[position[0]].source) {
                                showResult(options.columns[position[0]].source, str);
                            }
                        }, 500);
                    });
                    $(cell).html(editor);
                    $(cell).append(result);

                    // Current value
                    $(editor).focus();
                    $(editor).val('');

                    // Close editor handler
                    $(editor).blur(function () {
                        $(main).jexcel('closeEditor', $(cell), false);
                    });
                } else {
                    // Keep the current value
                    $(cell).addClass('edition');

                    var input = $(cell).find('input');

                    // Get content
                    if ($(input).length) {
                        var html = $(input).val();
                    } else {
                        var html = $(cell).html();
                    }

                    // Basic editor
                    var editor = document.createElement('input');
                    $(editor).prop('class', 'editor');
                    $(editor).css('width', $(cell).width());
                    $(cell).html(editor);

                    // Bind mask
                    if (options.columns[position[0]].mask) {
                        if (! $.fn.masked) {
                            console.error('Jexcel: it was not possible to load the mask plugin.');
                        } else {
                            $(editor).mask(options.columns[position[0]].mask, options.columns[position[0]].options)
                        }
                    }

                    // Current value
                    $(editor).focus();
                    if (! empty) {
                        $(editor).val(html);
                    }

                    // Close editor handler
                    $(editor).blur(function () {
                        $(main).jexcel('closeEditor', $(this).parent(), true);
                    });
                }
            }
        }
    },

    /**
     * Close the editor and save the information
     * 
     * @param object cell
     * @param boolean save
     * @return void
     */
    closeEditor : function(cell, save) {
        // Remove edition mode mark
        $(cell).removeClass('edition');

        // Id
        var id = $(this).prop('id');

        // Options
        var options = $.fn.jexcel.defaults[id];

        // Cell identification
        var position = $(cell).prop('id').split('-');

        // Get cell properties
        if (save == true) {
            // Before change
            if (typeof(options.columns[position[0]].onbeforechange) == 'function') {
                options.columns[position[0]].onbeforechange($(this), $(cell));
            }

            // If custom editor
            if (options.columns[position[0]].editor) {
                // Custom editor
                options.columns[position[0]].editor.closeEditor(cell, save);
            } else {
                // Native functions
                if (options.columns[position[0]].type == 'checkbox' || options.columns[position[0]].type == 'hidden') {
                    // Do nothing
                } else if (options.columns[position[0]].type == 'dropdown') {
                    // Get value
                    var value = $(cell).find('select').val();
                    var text = $(cell).find('select').find('option:selected').text();
                    // Set value
                    $(cell).html('<input type="hidden" value="' + value + '">' + text);
                } else if (options.columns[position[0]].type == 'autocomplete') {
                    // Set value
                    var obj = $(cell).find('li');
                    if (obj.length > 0) {
                        var value = $(cell).find('li').prop('id');
                        var text = $(cell).find('li').html();
                        $(cell).html('<input type="hidden" value="' + value + '">' + text);
                    } else {
                        $(cell).html('');
                    }
                } else if (options.columns[position[0]].type == 'calendar') {
                    var value = $(cell).find('.jcalendar_value').val();
                    var text = $(cell).find('.jcalendar_input').val();
                    $(cell).html('<input type="hidden" value="' + value + '">' + text);
                } else {
                    // Get content
                    var value = $(cell).find('.editor').val();
                    // For formulas
                    if (value.substr(0,1) == '=') {
                        if ($.fn.jexcel.defaults[id].dynamicColumns.indexOf($(cell).prop('id')) == -1) {
                            $.fn.jexcel.defaults[id].dynamicColumns.push($(cell).prop('id'));
                        }
                    }
                    $(cell).html(value);
                }
            }

            // Get value from column and set the default
            $.fn.jexcel.defaults[id].data[position[1]][position[0]] = $(this).jexcel('getValue', $(cell));

            // Change
            if (typeof(options.onchange) == 'function') {
                options.onchange($(this), $(cell), value);
            }

            // After changes
            $(this).jexcel('afterChange');
        } else {
            if (options.columns[position[0]].type == 'calendar') {
                // Do nothing - calendar will be closed without keeping the current value
            } else {
                // Restore value
                $(cell).html($.fn.jexcel.edition);
    
                // Finish temporary edition
                $.fn.jexcel.edition = null;
            }
        }
    },

    /**
     * Get the value from a cell
     * 
     * @param object cell
     * @return string value
     */
    getValue : function(cell) {
        var value = null;

        // If is a string get the cell object
        if (typeof(cell) != 'object') {
            // Convert in case name is excel liked ex. A10, BB92
            cell = $(this).jexcel('id', cell);
            // Get object based on a string ex. 12-1, 13-3
            cell = $(this).find('[id=' + cell +']');
        }

        // If column exists
        if ($(cell).length) {
            // Id
            var id = $(this).prop('id');

            // Global options
            var options = $.fn.jexcel.defaults[id];

            // Configuration
            var position = $(cell).prop('id').split('-');

            // Get value based on the type
            if (options.columns[position[0]].editor) {
                // Custom editor
                value = options.columns[position[0]].editor.getValue(cell);
            } else {
                // Native functions
                if (options.columns[position[0]].type == 'checkbox') {
                    // Get checkbox value
                    value = $(cell).find('input').is(':checked') ? 1 : 0;
                } else if (options.columns[position[0]].type == 'dropdown' || options.columns[position[0]].type == 'autocomplete' || options.columns[position[0]].type == 'calendar') {
                    // Get value
                    value = $(cell).find('input').val();
                } else if (options.columns[position[0]].type == 'currency') {
                    value = $(cell).html().replace( /\D/g, '');
                } else {
                    // Get default value
                    value = $(cell).find('input');
                    if ($(value).length) {
                        value = $(value).val(); 
                    } else {
                        value = $(cell).html();
                    }
                }
            }
        }

        return value;
    },

    /**
     * Set a cell value
     * 
     * @param object cell destination cell
     * @param object value value
     * @return void
     */
    setValue : function(cell, value, ignoreEvents) {
        // If is a string get the cell object
        if (typeof(cell) !== 'object') {
            // Convert in case name is excel liked ex. A10, BB92
            cell = $(this).jexcel('id', cell);
            // Get object based on a string ex. 12-1, 13-3
            cell = $(this).find('[id=' + cell +']');
        }

        // If column exists
        if ($(cell).length) {
            // Id
            var id = $(this).prop('id');

            // Main object
            var main = $(this);

            // Global options
            var options = $.fn.jexcel.defaults[id];

            // Go throw all cells
            $.each(cell, function(k, v) {
                // Cell identification
                var position = $(v).prop('id').split('-');

                // Before Change
                if (! ignoreEvents) {
                    if (typeof(options.columns[position[0]].onbeforechange) == 'function') {
                        options.columns[position[0]].onbeforechange($(this), $(v));
                    }
                }

                if (options.columns[position[0]].editor) {
                    // Custom editor
                    options.columns[position[0]].editor.setValue(v, value);
                } else if (options.columns[position[0]].readOnly == true) {
                    // Do nothing
                    value = null;
                } else {
                    // Native functions
                    if (options.columns[position[0]].type == 'checkbox') {
                        if (value == 1 || value == true) {
                            $(v).find('input').prop('checked', true);
                        } else {
                            $(v).find('input').prop('checked', false);
                        }
                    } else if (options.columns[position[0]].type == 'dropdown' || options.columns[position[0]].type == 'autocomplete') {
                        // Dropdown and autocompletes
                        key = '';
                        val = '';
                        if (value) {
                            if (options.columns[position[0]].combo[value]) {
                                key = value;
                                val = options.columns[position[0]].combo[value];
                            } else {
                                value = null;
                            }
                        }

                        $(v).html('<input type="hidden" value="' +  key + '">' + val);
                    } else if (options.columns[position[0]].type == 'calendar') {
                        val = '';
                        if (value) {
                            val = $.fn.jcalendar('label', value);
                        }
                        $(v).html('<input type="hidden" value="' + value + '">' + val);
                    } else {
                        if (value) {
                            if (value.substr(0,1) == '=') {
                                if ($.fn.jexcel.defaults[id].dynamicColumns.indexOf($(cell).prop('id')) == -1) {
                                    $.fn.jexcel.defaults[id].dynamicColumns.push($(cell).prop('id'));
                                }
                            }
                        }

                        $(v).html(value);
                    }
                }

                // Get value from column and set the default
                $.fn.jexcel.defaults[id].data[position[1]][position[0]] = value;

                // Change
                if (! ignoreEvents) {
                    if (typeof(options.onchange) == 'function') {
                        options.onchange($(this), $(v), value);
                    }
                }
            });

            // After changes
            if (! ignoreEvents) {
                $(this).jexcel('afterChange');
            }

            return true;
        } else {
            return false;
        }
    },

    /**
     * Update the cells selection
     * 
     * @param object o cell origin
     * @param object d cell destination
     * @return void
     */
    updateSelection : function(o, d) {
        // Main table
        var main = $(this);

        // Cells
        var cells = $(this).find('tbody td');
        var header = $(this).find('thead td');

        // Remove highlight
        $(cells).removeClass('highlight');
        $(cells).removeClass('highlight-left');
        $(cells).removeClass('highlight-right');
        $(cells).removeClass('highlight-top');
        $(cells).removeClass('highlight-bottom');

        // Update selected column
        $(header).removeClass('selected');
        $(cells).removeClass('selected');
        $(o).addClass('selected');

        // Define coordinates
        o = $(o).prop('id').split('-');
        d = $(d).prop('id').split('-');

        if (parseInt(o[0]) < parseInt(d[0])) {
            px = parseInt(o[0]);
            ux = parseInt(d[0]);
        } else {
            px = parseInt(d[0]);
            ux = parseInt(o[0]);
        }

        if (parseInt(o[1]) < parseInt(d[1])) {
            py = parseInt(o[1]);
            uy = parseInt(d[1]);
        } else {
            py = parseInt(d[1]);
            uy = parseInt(o[1]);
        }

        // Redefining styles
        for (i = px; i <= ux; i++) {
            for (j = py; j <= uy; j++) {
                $(this).find('#' + i + '-' + j).addClass('highlight');
                $(this).find('#' + px + '-' + j).addClass('highlight-left');
                $(this).find('#' + ux + '-' + j).addClass('highlight-right');
                $(this).find('#' + i + '-' + py).addClass('highlight-top');
                $(this).find('#' + i + '-' + uy).addClass('highlight-bottom');

                // Row and column headers
                $(main).find('#col-' + i).addClass('selected');
                $(main).find('#row-' + j).addClass('selected');
            }
        }

        // Find corner cell
        $(this).jexcel('updateCornerPosition');
    },

    /**
     * Update the cells move data TODO: copy multi columns - TODO!
     * 
     * @param object o cell origin
     * @param object d cell destination
     * @return void
     */
    updateCornerSelection : function(current) {
        // Main table
        var main = $(this);

        // Remove selection
        var cells = $(this).find('tbody td');
        $(cells).removeClass('selection');
        $(cells).removeClass('selection-left');
        $(cells).removeClass('selection-right');
        $(cells).removeClass('selection-top');
        $(cells).removeClass('selection-bottom');

        // Get selection
        var selection = $(this).find('tbody td.highlight');

        // Get elements first and last
        var s = $(selection[0]).prop('id').split('-');
        var d = $(selection[selection.length - 1]).prop('id').split('-');

        // Get current
        var c = $(current).prop('id').split('-');

        // Vertical copy
        if (c[1] > d[1] || c[1] < s[1]) {
            // Vertical
            var px = parseInt(s[0]);
            var ux = parseInt(d[0]);
            if (parseInt(c[1]) > parseInt(d[1])) {
                var py = parseInt(d[1]) + 1;
                var uy = parseInt(c[1]);
            } else {
                var py = parseInt(c[1]);
                var uy = parseInt(s[1]) - 1;
            }
        } else if (c[0] > d[0] || c[0] < s[0]) {
            // Horizontal copy
            var py = parseInt(s[1]);
            var uy = parseInt(d[1]);
            if (parseInt(c[0]) > parseInt(d[0])) {
                var px = parseInt(d[0]) + 1;
                var ux = parseInt(c[0]);
            } else {
                var px = parseInt(c[0]);
                var ux = parseInt(s[0]) - 1;
            }
        }

        for (j = py; j <= uy; j++) {
            for (i = px; i <= ux; i++) {
                $(this).find('#' + i + '-' + j).addClass('selection');
                $(this).find('#' + i + '-' + py).addClass('selection-top');
                $(this).find('#' + i + '-' + uy).addClass('selection-bottom');
                $(this).find('#' + px + '-' + j).addClass('selection-left');
                $(this).find('#' + ux + '-' + j).addClass('selection-right');
            }
        }

        //$(this).jexcel('updateCornerPosition');
    },

    /**
     * Update corner position
     * 
     * @return void
     */
    updateCornerPosition : function() {
        var cells = $(this).find('.highlight');
        if ($(cells).length) { 
            corner = $(cells).last();

            // Get the position of the corner helper
            var t = parseInt($(corner).offset().top) + $(corner).height() + 5;
            var l = parseInt($(corner).offset().left) + $(corner).width() + 5;

            // Place the corner in the correct place
            $('.jexcel_corner').css('top', t);
            $('.jexcel_corner').css('left', l);
        }
    },

    /**
     * Get the data from a row
     * 
     * @param integer row number
     * @return string value
     */
    getRowData : function(row) {
       // Get row
       row = $(this).find('#row-' + row).parent().find('td').not(':first');

       // String
       var str = '';

       // Search all tds in a row
       if (row.length > 0) {
          for (i = 0; i < row.length; i++) {
             str += $(this).jexcel('getValue', $(row)[i]) + ',';
          }
       }

       return str;
    },

    /**
     * Get the whole table data
     * 
     * @param integer row number
     * @return string value
     */
    getData : function(highlighted) {
        // Control vars
        var dataset = [];
        var px = 0;
        var py = 0;

        // Column and row length
        var x = $(this).find('thead tr td').not(':first').length;
        var y = $(this).find('tbody tr').length;

        // Go through the columns to get the data
        for (j = 0; j < y; j++) {
            px = 0;
            for (i = 0; i < x; i++) {
                // Cell
                cell = $(this).find('#' + i + '-' + j);

                // Cell selected or fullset
                if (! highlighted || $(cell).hasClass('highlight')) {
                    // Get value
                    if (! dataset[py]) {
                        dataset[py] = [];
                    }
                    dataset[py][px] = $(this).jexcel('getValue', $(cell));
                    px++;
                }
            }
            if (px > 0) {
                py++;
            }
        }

       return dataset;
    },

    /**
     * Copy method
     * 
     * @param bool highlighted - Get only highlighted cells
     * @param delimiter - \t default to keep compatibility with excel
     * @return string value
     */
    copy : function(highlighted, delimiter, returnData) {
        if (! delimiter) {
            delimiter = "\t";
        }

        var str = '';
        var row = '';
        var val = '';
        var pc = false;
        var pr = false;

        // Column and row length
        var x = $(this).find('thead tr td').not(':first').length;
        var y = $(this).find('tbody tr').length;

        // Go through the columns to get the data
        for (j = 0; j < y; j++) {
            row = '';
            pc = false;
            for (i = 0; i < x; i++) {
                // Get cell
                cell = $(this).find('#' + i + '-' + j);

                // If cell is highlighted
                if (! highlighted || $(cell).hasClass('highlight')) {
                    if (pc) {
                        row += delimiter;
                    }
                    // Get value
                    val = $(this).jexcel('getValue', $(cell));
                    if (val.match(/,/g)) {
                        val = '"' + val + '"'; 
                    }
                    row += val;
                    pc = true;
                }
            }
            if (row) {
                if (pr) {
                    str += "\n";
                }
                str += row;
                pr = true;
            }
        }

        // Create a hidden textarea to copy the values
        if (! returnData) {
            txt = $('.jexcel_textarea');
            $(txt).val(str);
            $(txt).select();
            document.execCommand("copy");
        }

        return str;
    },

    /**
     * Paste method TODO: if the clipboard is larger than the table create automatically columns/rows?
     * 
     * @param integer row number
     * @return string value
     */
    paste : function(cell, data) {
        // Id
        var id = $(this).prop('id');

        // Data
        data = data.split("\r\n");

        // Initial position
        var position = $(cell).prop('id');
        if (position) {
            position = position.split('-');
            var x = position[0];
            var y = position[1];
    
            // Automatic adding new rows when the copied data is larger then the table
            if (parseInt(y + data.length) > $.fn.jexcel.defaults[id].data.length) {
                $(this).jexcel('insertRow', null, parseInt(y) + data.length - $.fn.jexcel.defaults[id].data.length);
            }
    
            // Go through the columns to get the data
            for (j = 0; j < data.length; j++) {
                // Explode column values
                row = data[j].split("\t");
                for (i = 0; i < row.length; i++) {
                    // Get cell
                    cell = $(this).find('#' + (parseInt(i) + parseInt(x))  + '-' + (parseInt(j) + parseInt(y)));
    
                    // If cell exists
                    if ($(cell).length > 0) {
                        $(this).jexcel('setValue', $(cell), row[i]);
                    }
                }
            }
        }
    },

    /**
     * TODO: Insert a new column
     * 
     * @param  object properties
     * @return void
     */
    insertColumn : function (properties) {
        // Id
        var id = $(this).prop('id');

        // Main configuration
        var options = $.fn.jexcel.defaults[id];

        // Current number of columns
        var num = options.colHeaders.length;

        // Adding the headers
        options.colHeaders[num] = 'test';

        // Default header cell properties
        width = options.colWidths[num] || 50;
        align = options.colAlignments[num] || 'left';

        // Other column types
        var td =  '<td id="col-' + num + '" width="' + width + '" align="' + align +'" title="' + options.colHeaders[i] + '">' + options.colHeaders[i] + '</td>';

        // Add element to the table
        //var tr = $(this).find('thead.label tr')[0];
        //$(tr).append(td);
    },

    /**
     * Insert a new row TODO: add relative row
     * 
     * @param object relativeRow - add new row from line number, or null for the end of the table
     * @param object numLines - how many lines to be included
     * 
     * @return void
     */
    insertRow : function(relativeRow, numLines) {
        // Id
        var id = $(this).prop('id');

        // Main configuration
        var options = $.fn.jexcel.defaults[id];

        // Num lines
        if (! numLines) {
            // Add one line is the default
            numLines = 1;
        } else if (numLines > 100) {
            // TODO: is this a good practise to limit the user will?
            numLines = 100
        } 

        j = parseInt($.fn.jexcel.defaults[id].data.length);

        // Adding lines
        for (row = 0; row < numLines; row++) {
            // New row
            var tr = '<td id="row-' + j + '" class="label">' + (j + 1) + '</td>';

            // New data
            $.fn.jexcel.defaults[id].data[j] = [];

            for (i = 0; i < $.fn.jexcel.defaults[id].colHeaders.length; i++) {
                // New Data
                $.fn.jexcel.defaults[id].data[j][i] = '';

                // Aligment
                align = $.fn.jexcel.defaults[id].colAlignments[i] || 'left';

                // Hidden column
                if ($.fn.jexcel.defaults[id].columns[i].type == 'hidden') {
                    tr += '<td id="' + i + '-' + j + '" style="display:none;"></td>';
                } else {
                    // Native options
                    if ($.fn.jexcel.defaults[id].columns[i].type == 'checkbox') {
                        contentCell = '<input type="checkbox">';
                    } else if ($.fn.jexcel.defaults[id].columns[i].type == 'dropdown' || $.fn.jexcel.defaults[id].columns[i].type == 'autocomplete' || $.fn.jexcel.defaults[id].columns[i].type == 'calendar') {
                        contentCell = '<input type="hidden" value="">';
                    } else {
                        contentCell = '';
                    }

                    tr += '<td id="' + i + '-' + j + '" align="' + align +'">' + contentCell + '</td>';
                }
            }

            tr = '<tr>' + tr + '</tr>';

            $(this).find('tbody').append(tr);

            j++;
        }
    },

    /**
     * Update column source for dropboxes
     */
    setSource : function (column, source) {
        // In case the column is an object
        if (typeof(column) == 'object') {
            column = $(column).prop('id').split('-');
            column = column[0];
        }

        // Id
        var id = $(this).prop('id');

        // Update defaults
        $.fn.jexcel.defaults[id].columns[column].source = source;
        $.fn.jexcel.defaults[id].columns[column].combo = $(this).jexcel('createCombo', source);
    },

    /**
     * After change
     */
    afterChange : function() {
        // Id
        var id = $(this).prop('id');

        // Dynamic updates
        if ($.fn.jexcel.defaults[id].dynamicColumns.length > 0) {
            $(this).jexcel('formula');
        }

        // After Changes
        if (typeof($.fn.jexcel.defaults[id].onafterchange) == 'function') {
            $.fn.jexcel.defaults[id].onafterchange($(this));
        }

        // Update settings
        $(this).jexcel('updateSettings');
    },

    /**
     * Helper function to copy data using the corner icon
     */
    copyData : function(o, d) {
        var data = $(this).jexcel('getData', true);

        // Cells
        var px = parseInt(o[0]);
        var ux = parseInt(d[0]);
        var py = parseInt(o[1]);
        var uy = parseInt(d[1]);

        // Copy data procedure
        var posx = 0;
        var posy = 0;
        for (j = py; j <= uy; j++) {
            // Controls
            if (data[posy] == undefined) {
                posy = 0;
            }
            posx = 0;

            // Data columns
            for (i = px; i <= ux; i++) {
                // Column
                if (data[posy] == undefined) {
                    posx = 0;
                } else if (data[posy][posx] == undefined) {
                    posx = 0;
                }

                // Get cell
                cell = $(this).find('#' + i + '-' + j);

                // Update non-readonly
                if (! $(cell).hasClass('readonly')) {
                    $(this).jexcel('setValue', cell, data[posy][posx]);
                }
                posx++;
            }
            posy++;
        }
    },

    /**
     * Apply formula to a column
     */
    formula : function() {
        // Keep instannce of this object
        var main = $(this);

        // Id
        var id = $(this).prop('id');

        // Custom formulas
        if ($.fn.jexcel.defaults[id].formulas) {
            var formulas = $.fn.jexcel.defaults[id].formulas;
            // Set instance
            $.fn.jexcel.defaults[id].formulas.instance = this;
        }

        // Dynamic columns
        var columns = $.fn.jexcel.defaults[id].dynamicColumns;

        // Process columns
        $.each(columns, function (k, column) {
            // Get value from the column
            formula = $(main).jexcel('getValue', column);

            // Column value is a formula
            if (formula) {
                if (formula.substr(0,1) == '=') {
                    // Get method name  TODO: Get a simple formula such as
                    fnc = /=([a-zA-Z]+)\(/.exec(formula);
                    if (fnc) {
                        fnc = fnc[1].toLowerCase();
                        // Get arguments from string
                        arg = /\((.*)\)/.exec(formula);
                        // Call method
                        if (typeof(formulas[fnc]) == 'function') {
                            value = formulas[fnc](arg[1]);
                        } else {
                            value = formula;
                            console.error('Jexcel: formula method not found: ' + fnc);
                        }
                    } else {
                        // Value format
                        value = $(main).jexcel('basicFormula', formula);
                    }

                    // Set value
                    if (value === null || isNaN(value)) {
                        $(main).find('#' + column).addClass('error');
                        value = '<input type="hidden" value="' + formula + '">#ERROR';
                        // Update cell content
                        $(main).find('#' + column).html(value);
                    } else {
                        $(main).find('#' + column).removeClass('error');
                        value = '<input type="hidden" value="' + formula + '">' + value;
                        // Update cell content
                        $(main).find('#' + column).html(value);
                    }
                } else {
                    // Remove any existing calculation error
                    $(main).find('#' + column).removeClass('error');
                    // No longer dynamic
                    columns.splice(k, 1);
                }
            } else {
                // Remove any existing calculation error
                $(main).find('#' + column).removeClass('error');
                // No longer dynamic
                columns.splice(k, 1);
            }
        });
    },

    /**
     * Basic formula without methods
     *
     * @param string formula
     * @return total
     */
    basicFormula : function (formula) {
        // Remove =
        formula = formula.substr(1);
        // Get variables from the formula
        var exp = formula.match(/[A-Z][0-9]+/g);
        // Replace variables by value on the formula
        for (var i = 0; i < exp.length; i++) {
            id = $.fn.jexcel('id', exp[i]).split('-');
            formula = formula.replace(exp[i], $(this).jexcel('getValue', exp[i]));
        }

        try {
            return eval(formula);
        } catch (e) {
            return null;
        }
    },

    // Combo
    createCombo : function (result) {
        // Creating the mapping
        var combo = [];
        if (result.length > 0) {
            for (var j = 0; j < result.length; j++) {
                if (typeof(result[j]) == 'object') {
                    key = result[j].id
                    val = result[j].name;
                } else {
                    key = result[j];
                    val = result[j];
                }
                combo[key] = val;
            }
        }

        return combo;
    },

    /**
     * Multi-utility helper
     * 
     * @param object options { action: METHOD_NAME }
     * @return mixed
     */
    helper : function (options) {
        var data = [];
        if (typeof(options) == 'object') {
            // Return a empty bidimensional array
            if (options.action == 'createEmptyData') {
                var x = options.cols || 10;
                var y = options.rows || 100;
                for (j = 0; j < y; j++) {
                    data[j] = [];
                    for (i = 0; i < x; i++) {
                        data[j][i] = '';
                    }
                }
            }
        }

        return data;
    },

    /**
     * Convert excel like column to jexcel id
     * 
     * @param string id
     * @return string id
     */
    id : function (id) {
        var t = /^[a-zA-Z]+/.exec(id);
        if (t) {
            var code = 0;
            for (var i = 0; i < t[0].length; i++) {
                code += parseInt(t[0].charCodeAt(i) - 65);
            }
            id = code + '-' + (parseInt(/[0-9]+$/.exec(id)) - 1);
        }
        return id;
    },

    /**
     * Download CSV table TODO: improve chartset
     * 
     * @return null
     */
    download : function () {
        //if ($.csv) {
        //    var data = $(this).jexcel('getData', false);
        //    data = $.csv.fromArrays(data);
        //} else {
            var data = $(this).jexcel('copy', false, ',', true);
        //}

        var pom = document.createElement('a');
        var blob = new Blob([data], {type: 'text/csv;charset=utf-8;'});
        var url = URL.createObjectURL(blob);
        pom.href = url;
        pom.setAttribute('download', 'jexcelTable.csv');
        pom.click();
    }
};

$.fn.jexcel = function( method ) {
    if ( methods[method] ) {
        return methods[ method ].apply( this, Array.prototype.slice.call( arguments, 1 ));
    } else if ( typeof method === 'object' || ! method ) {
        return methods.init.apply( this, arguments );
    } else {
        $.error( 'Method ' +  method + ' does not exist on jQuery.tooltip' );
    }
};

})( jQuery );

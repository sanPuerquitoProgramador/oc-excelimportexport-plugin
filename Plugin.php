<?php namespace WRvE\ExcelImportExport;

use Backend;
use Event;
use System\Classes\PluginBase;

/**
 * ExcelImportExport Plugin Information File
 */
class Plugin extends PluginBase
{
    /**
     * Returns information about this plugin.
     *
     * @return array
     */
    public function pluginDetails()
    {
        return [
            'name' => 'Excel Import & Export',
            'description' => 'Adds support for Excel to ImportExportController',
            'author' => 'Ruben van Erk',
            'icon' => 'icon-file-excel-o',
        ];
    }

    /**
     * Boot method, called right before the request route.
     *
     * @return array
     */
    public function boot()
    {
        Event::listen('backend.form.extendFields', function (Backend\Widgets\Form $widget) {
            if (!$widget->getController()->implement) {
                return;
            }

            if (!in_array('WRvE\ExcelImportExport\Behaviors\ExcelImportExportController', $widget->getController()->implement)) {
                return;
            }

            if (!$widget->model instanceof Backend\Models\ImportModel) {
                return;
            }

            if (!$widget->getField('import_file')) {
                return;
            }

            $widget->addFields([
                'import_file' => [
                    'label' => 'Importar datos',
                    'type' => 'fileupload',
                    'mode' => 'file',
                    'span' => 'left',
                    'fileTypes' => 'csv,xlsx',
                    'useCaption' => false,
                ]
            ]);
        });
    }
}

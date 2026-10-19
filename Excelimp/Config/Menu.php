<?php

/**
 * 菜单配置
 * 挂载到 内容 -> 内容管理 分组下
 */

return [

    'admin' => [

        'content' => [

            'left' => [

                'content-module' => [

                    'link' => [

                        'app-excelimp' => [
                            'name' => 'Excel数据导入',
                            'icon' => 'fa fa-file-excel-o',
                            'uri' => 'excelimp/home/index',
                        ],

                    ]
                ],

            ],

        ],

    ],

];

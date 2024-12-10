<!--
* @Description  测试docx导出
* @FileName  App
* @Date 2024-11-22 14:45:27
!-->
<template>
  <div id="app">
    <input type="file" @input="handleFileUpload" />
    <br />
    <button v-if="show" @click="exportDocx">HTML导出docx</button>
    <!-- <button v-if="show" @click="jsonExportDocx">JSON导出docx</button> -->
    <!-- <button @click="exportDocx2">输出默认文本</button> -->
  </div>
</template>
<script setup lang="js">
import { onBeforeMount, ref } from 'vue';
import { Html2Docx } from '@/utils/Html2Docx/index.ts';
import { Json2Docx } from '@/utils/Json2Docx/index.ts';
import { ToTextBox } from '@/utils/ToTextBox/index.ts';
import { openDocxAsZip, getBookmarkStyleMap } from '@/utils/readDocx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
const jsZip = new JSZip();
const json = {
    "type": "doc",
    "content": [
        {
            "type": "heading",
            "attrs": {
                "gzTextAlign": "center",
                "gzSpacing": {
                    "lineHeight": "2.4",
                    "textIndent": "",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "17pt",
                    "marginBottom": "16.5pt"
                },
                "level": 1
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "22pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "财务简历"
                }
            ]
        },
        {
            "type": "heading",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.75",
                    "textIndent": "",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "13pt",
                    "marginBottom": "13pt"
                },
                "level": 2
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "黑体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "16pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "个人信息"
                }
            ]
        },
        {
            "type": "paragraph",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.5",
                    "textIndent": "",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "",
                    "marginBottom": ""
                }
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "姓名"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "： XXX"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "联系方式"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "： 手机：XXX，邮箱：XXX"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "现居住地"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "： XXX"
                }
            ]
        },
        {
            "type": "heading",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.75",
                    "textIndent": "",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "13pt",
                    "marginBottom": "13pt"
                },
                "level": 2
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "黑体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "16pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "教育背景"
                }
            ]
        },
        {
            "type": "paragraph",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.5",
                    "textIndent": "2em",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "",
                    "marginBottom": ""
                }
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "XXX大学"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": " "
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "italic"
                        }
                    ],
                    "text": "XXXX年XX月 - XXXX年XX月"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "  "
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "财务管理专业"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": " "
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "本科"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": " "
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "学位证书"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "、毕业证书"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "  "
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": "#ffff00"
                            }
                        }
                    ],
                    "text": "在校期间，系统学习了财务管理、会计学、金融市场等核心课程，积极参与各类财务案例分析，培养了扎实的理论基础和实践应用能力。"
                }
            ]
        },
        {
            "type": "heading",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.75",
                    "textIndent": "",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "13pt",
                    "marginBottom": "13pt"
                },
                "level": 2
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "黑体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "16pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "工作经历\t"
                }
            ]
        },
        {
            "type": "paragraph",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.5",
                    "textIndent": "2em",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "",
                    "marginBottom": ""
                }
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "XXX公司 财务部"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": " 财务专员 XXXX年XX月 - 至今"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "  负责公司日常财务事务处理，包括凭证录入、账目核对、报表编制等工作。熟练掌握财务软件，能够高效完成各项财务数据分析任务。参与年度财务预算的编制与执行情况分析，为公司决策提供有力支持。期间，成功协助部门完成了多次内部审计和外部审计工作，得到了领导和同事的认可。"
                }
            ]
        },
        {
            "type": "heading",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.75",
                    "textIndent": "",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "13pt",
                    "marginBottom": "13pt"
                },
                "level": 2
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "黑体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "16pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "项目经历"
                }
            ]
        },
        {
            "type": "paragraph",
            "attrs": {
                "gzTextAlign": "distribute",
                "gzSpacing": {
                    "lineHeight": "1.5",
                    "textIndent": "2em",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "",
                    "marginBottom": ""
                }
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        },
                        {
                            "type": "bold"
                        }
                    ],
                    "text": "XXX项目"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": " 项目财务助理 XXXX年XX月 - XXXX年XX月"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "  在项目中担任财务助理角色，负责项目预算的编制、成本控制以及财务结算工作。通过精细化的财务管理，确保项目各项费用合理支出，有效控制了项目成本。积极与项目团队成员沟通协作，共同推进项目进度，最终实现了项目按时交付和盈利目标。"
                }
            ]
        },
        {
            "type": "heading",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.75",
                    "textIndent": "",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "13pt",
                    "marginBottom": "13pt"
                },
                "level": 2
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "黑体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "16pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "专业技能与证书"
                }
            ]
        },
        {
            "type": "paragraph",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.5",
                    "textIndent": "2em",
                    "marginLeft": "1em",
                    "marginRight": "",
                    "marginTop": "",
                    "marginBottom": ""
                }
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "1. "
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "#ff0000",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "熟练掌握"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "财务软件（如：金蝶、用友等），能够高效处理各类财务数据；"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "  2. "
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "#ff0000",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "熟悉"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "国家财经法规及税收政策，具备良好的税务筹划能力；"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "  3. 持有会计从业资格证书，具备扎实的会计理论基础和实务操作能力；"
                },
                {
                    "type": "hardBreak",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ]
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "  4. 通过英语四级和六级考试，具备良好的英语听说读写能力，能够阅读英文财务报表及相关资料。"
                }
            ]
        },
        {
            "type": "heading",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "1.75",
                    "textIndent": "",
                    "marginLeft": "0em",
                    "marginRight": "",
                    "marginTop": "13pt",
                    "marginBottom": "13pt"
                },
                "level": 2
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "黑体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "16pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "自我评价与职业规划"
                }
            ]
        },
        {
            "type": "paragraph",
            "attrs": {
                "gzTextAlign": "justify",
                "gzSpacing": {
                    "lineHeight": "2",
                    "textIndent": "2em",
                    "marginLeft": "",
                    "marginRight": "",
                    "marginTop": "",
                    "marginBottom": ""
                }
            },
            "content": [
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "本人性格"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": "#ffff00"
                            }
                        }
                    ],
                    "text": "开朗、稳重、有活力，待人热情、真诚；工作认真负责"
                },
                {
                    "type": "text",
                    "marks": [
                        {
                            "type": "textStyle",
                            "attrs": {
                                "fontFamily": "宋体",
                                "color": "rgb(0, 0, 0)",
                                "gzFontSize": "12pt",
                                "gzHighlight": ""
                            }
                        }
                    ],
                    "text": "，积极主动，能吃苦耐劳，勇于承受压力，勇于创新；有很强的组织能力和团队协作精神，具有较强的适应能力；纪律性强，工作积极配合；意志坚强，具有较强的无私奉献精神。在未来的职业发展中，我希望能够继续深耕财务管理领域，不断提升自己的专业素养和综合能力。通过参加专业培训、考取相关证书等方式不断丰富自己的知识体系和实践经验。我也期待有机会挑战更高层次的财务管理岗位，为公司创造更大的价值并实现个人职业生涯的跨越式发展。"
                }
            ]
        }
    ]
}
const html =
  `<h1 style="text-align: center; line-height: 2.4; margin-top: 17pt; margin-bottom: 16.5pt"><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 22pt">财务简历</span></h1><h2 style="text-align: justify; line-height: 1.75; margin-top: 13pt; margin-bottom: 13pt"><span style="font-family: 黑体; color: rgb(0, 0, 0); font-size: 16pt">个人信息</span></h2><p style="text-align: justify; line-height: 1.5"><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt"><strong>姓名</strong>： XXX<br><strong>联系方式</strong>： 手机：XXX，邮箱：XXX<br><strong>现居住地</strong>： XXX</span></p><h2 style="text-align: justify; line-height: 1.75; margin-top: 13pt; margin-bottom: 13pt"><span style="font-family: 黑体; color: rgb(0, 0, 0); font-size: 16pt">教育背景</span></h2><p style="text-align: justify; line-height: 1.5; text-indent: 2em"><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt"><strong>XXX大学</strong> <em>XXXX年XX月 - XXXX年XX月</em><br>  <strong>财务管理专业</strong> <strong>本科</strong> <strong>学位证书</strong>、毕业证书<br>  </span><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt; background-color: #ffff00">在校期间，系统学习了财务管理、会计学、金融市场等核心课程，积极参与各类财务案例分析，培养了扎实的理论基础和实践应用能力。</span></p><h2 style="text-align: justify; line-height: 1.75; margin-top: 13pt; margin-bottom: 13pt"><span style="font-family: 黑体; color: rgb(0, 0, 0); font-size: 16pt">工作经历	</span></h2><p style="text-align: justify; line-height: 1.5; text-indent: 2em"><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt"><strong>XXX公司 财务部</strong> 财务专员 XXXX年XX月 - 至今<br>  负责公司日常财务事务处理，包括凭证录入、账目核对、报表编制等工作。熟练掌握财务软件，能够高效完成各项财务数据分析任务。参与年度财务预算的编制与执行情况分析，为公司决策提供有力支持。期间，成功协助部门完成了多次内部审计和外部审计工作，得到了领导和同事的认可。</span></p><h2 style="text-align: justify; line-height: 1.75; margin-top: 13pt; margin-bottom: 13pt"><span style="font-family: 黑体; color: rgb(0, 0, 0); font-size: 16pt">项目经历</span></h2><p style="text-align: justify; text-align-last: justify; line-height: 1.5; text-indent: 2em"><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt"><strong>XXX项目</strong> 项目财务助理 XXXX年XX月 - XXXX年XX月<br>  在项目中担任财务助理角色，负责项目预算的编制、成本控制以及财务结算工作。通过精细化的财务管理，确保项目各项费用合理支出，有效控制了项目成本。积极与项目团队成员沟通协作，共同推进项目进度，最终实现了项目按时交付和盈利目标。</span></p><h2 style="text-align: justify; line-height: 1.75; margin-top: 13pt; margin-bottom: 13pt"><span style="font-family: 黑体; color: rgb(0, 0, 0); font-size: 16pt">专业技能与证书</span></h2><p style="text-align: justify; line-height: 1.5; text-indent: 2em; margin-left: 1em"><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt">1. </span><span style="font-family: 宋体; color: #ff0000; font-size: 12pt">熟练掌握</span><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt">财务软件（如：金蝶、用友等），能够高效处理各类财务数据；<br>  2. </span><span style="font-family: 宋体; color: #ff0000; font-size: 12pt">熟悉</span><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt">国家财经法规及税收政策，具备良好的税务筹划能力；<br>  3. 持有会计从业资格证书，具备扎实的会计理论基础和实务操作能力；<br>  4. 通过英语四级和六级考试，具备良好的英语听说读写能力，能够阅读英文财务报表及相关资料。</span></p><h2 style="text-align: justify; line-height: 1.75; margin-left: 0em; margin-top: 13pt; margin-bottom: 13pt"><span style="font-family: 黑体; color: rgb(0, 0, 0); font-size: 16pt">自我评价与职业规划</span></h2><p style="text-align: justify; line-height: 2; text-indent: 2em"><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt">本人性格</span><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt; background-color: #ffff00">开朗、稳重、有活力，待人热情、真诚；工作认真负责</span><span style="font-family: 宋体; color: rgb(0, 0, 0); font-size: 12pt">，积极主动，能吃苦耐劳，勇于承受压力，勇于创新；有很强的组织能力和团队协作精神，具有较强的适应能力；纪律性强，工作积极配合；意志坚强，具有较强的无私奉献精神。在未来的职业发展中，我希望能够继续深耕财务管理领域，不断提升自己的专业素养和综合能力。通过参加专业培训、考取相关证书等方式不断丰富自己的知识体系和实践经验。我也期待有机会挑战更高层次的财务管理岗位，为公司创造更大的价值并实现个人职业生涯的跨越式发展。</span></p>`;

const show = ref(false);
const templateDocx = ref(null);
const exportDocx = async () => {
  const xmlMap = await openDocxAsZip(templateDocx.value);
  const styleMap = await getBookmarkStyleMap(xmlMap);
  const o = new Html2Docx(html, {
  });

  const blob = await o.exportDocx();
  saveAs(blob, 'test.docx');
};
/**
 * 方案修改，html使用时存在问题
 */
const jsonExportDocx = async () => {
  const xmlMap = await openDocxAsZip(templateDocx.value);
  const styleMap = await getBookmarkStyleMap(xmlMap);
  const o = new Json2Docx(json, {
    styleMap
  });

  const blob = await o.exportDocx();
  saveAs(blob, 'test.docx');
};
const handleFileUpload = async (event) => {
  const file = event.target.files[0]; // 获取用户选择的文件
  templateDocx.value = await jsZip.loadAsync(file);
  show.value = true;
};
const exportDocx2 = async () => {

  const o = new ToTextBox(html);
  const blob = await o.exportDocx();
  saveAs(blob, 'test.docx');
};
</script>
<style scoped></style>

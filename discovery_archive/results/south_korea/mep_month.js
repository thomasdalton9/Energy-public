









            var c1,c2,c3,c4,c5,c6,c8,c9;
            var count = $("#selCount option:selected").val();
            var point = "";
            if(count == 0) point = "";
            else if(count == 1) point = ".0";
            else if(count == 2) point = ".00";
            else if(count == 3) point = ".000";
            gridData = [];

            
                c1 = textFormmat("161478.0",count);
                c3 = textFormmat("96513.0",count);
                c4 = textFormmat("71992.0",count);
                c5 = textFormmat("24521.0",count);
                c6 = textFormmat("34.1",count);
                c7 = textFormmat("2026/10/01(19:00)",count);
                c8 = ("54935.0" == 0) ? "-" : Math.round("54935.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/10/01(04:00)" == "") ? "-" : textFormmat("2026/10/01(04:00)",count);

                gridData.push({"year":"2026",
                                "month":"10",
                                "day":"01",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("160952.0",count);
                c3 = textFormmat("103056.0",count);
                c4 = textFormmat("85902.0",count);
                c5 = textFormmat("17154.0",count);
                c6 = textFormmat("20",count);
                c7 = textFormmat("2026/09/02(17:00)",count);
                c8 = ("61582.0" == 0) ? "-" : Math.round("61582.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/09/02(04:00)" == "") ? "-" : textFormmat("2026/09/02(04:00)",count);

                gridData.push({"year":"2026",
                                "month":"09",
                                "day":"02",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("160411.0",count);
                c3 = textFormmat("103589.0",count);
                c4 = textFormmat("95321.0",count);
                c5 = textFormmat("8268.0",count);
                c6 = textFormmat("8.7",count);
                c7 = textFormmat("2026/08/07(19:00)",count);
                c8 = ("64455.0" == 0) ? "-" : Math.round("64455.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/08/07(05:00)" == "") ? "-" : textFormmat("2026/08/07(05:00)",count);

                gridData.push({"year":"2026",
                                "month":"08",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("159432.0",count);
                c3 = textFormmat("103519.0",count);
                c4 = textFormmat("92914.0",count);
                c5 = textFormmat("10605.0",count);
                c6 = textFormmat("11.4",count);
                c7 = textFormmat("2026/07/13(19:00)",count);
                c8 = ("57711.0" == 0) ? "-" : Math.round("57711.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/07/13(04:00)" == "") ? "-" : textFormmat("2026/07/13(04:00)",count);

                gridData.push({"year":"2026",
                                "month":"07",
                                "day":"13",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("159148.0",count);
                c3 = textFormmat("92978.0",count);
                c4 = textFormmat("82193.0",count);
                c5 = textFormmat("10785.0",count);
                c6 = textFormmat("13.1",count);
                c7 = textFormmat("2026/06/19(15:00)",count);
                c8 = ("57870.0" == 0) ? "-" : Math.round("57870.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/06/19(05:00)" == "") ? "-" : textFormmat("2026/06/19(05:00)",count);

                gridData.push({"year":"2026",
                                "month":"06",
                                "day":"19",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("159088.0",count);
                c3 = textFormmat("84327.0",count);
                c4 = textFormmat("76629.0",count);
                c5 = textFormmat("7698.0",count);
                c6 = textFormmat("10",count);
                c7 = textFormmat("2026/05/26(17:00)",count);
                c8 = ("51393.0" == 0) ? "-" : Math.round("51393.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/05/26(04:00)" == "") ? "-" : textFormmat("2026/05/26(04:00)",count);

                gridData.push({"year":"2026",
                                "month":"05",
                                "day":"26",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("158337.0",count);
                c3 = textFormmat("80768.0",count);
                c4 = textFormmat("73633.0",count);
                c5 = textFormmat("7135.0",count);
                c6 = textFormmat("9.7",count);
                c7 = textFormmat("2026/04/09(15:00)",count);
                c8 = ("55015.0" == 0) ? "-" : Math.round("55015.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/04/09(04:00)" == "") ? "-" : textFormmat("2026/04/09(04:00)",count);

                gridData.push({"year":"2026",
                                "month":"04",
                                "day":"09",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("158244.0",count);
                c3 = textFormmat("88072.0",count);
                c4 = textFormmat("78424.0",count);
                c5 = textFormmat("9648.0",count);
                c6 = textFormmat("12.3",count);
                c7 = textFormmat("2026/03/18(11:00)",count);
                c8 = ("55873.0" == 0) ? "-" : Math.round("55873.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/03/18(04:00)" == "") ? "-" : textFormmat("2026/03/18(04:00)",count);

                gridData.push({"year":"2026",
                                "month":"03",
                                "day":"18",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("157285.0",count);
                c3 = textFormmat("100026.0",count);
                c4 = textFormmat("88950.0",count);
                c5 = textFormmat("11076.0",count);
                c6 = textFormmat("12.5",count);
                c7 = textFormmat("2026/02/10(10:00)",count);
                c8 = ("63804.0" == 0) ? "-" : Math.round("63804.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/02/10(04:00)" == "") ? "-" : textFormmat("2026/02/10(04:00)",count);

                gridData.push({"year":"2026",
                                "month":"02",
                                "day":"10",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("157113.0",count);
                c3 = textFormmat("104185.0",count);
                c4 = textFormmat("88548.0",count);
                c5 = textFormmat("15637.0",count);
                c6 = textFormmat("17.7",count);
                c7 = textFormmat("2026/01/22(09:00)",count);
                c8 = ("67283.0" == 0) ? "-" : Math.round("67283.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/01/22(04:00)" == "") ? "-" : textFormmat("2026/01/22(04:00)",count);

                gridData.push({"year":"2026",
                                "month":"01",
                                "day":"22",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("156637.0",count);
                c3 = textFormmat("102715.0",count);
                c4 = textFormmat("82812.0",count);
                c5 = textFormmat("19903.0",count);
                c6 = textFormmat("24",count);
                c7 = textFormmat("2025/12/23(10:00)",count);
                c8 = ("60732.0" == 0) ? "-" : Math.round("60732.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2025/12/23(04:00)" == "") ? "-" : textFormmat("2025/12/23(04:00)",count);

                gridData.push({"year":"2025",
                                "month":"12",
                                "day":"23",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("156253.0",count);
                c3 = textFormmat("95666.0",count);
                c4 = textFormmat("76163.0",count);
                c5 = textFormmat("19503.0",count);
                c6 = textFormmat("25.6",count);
                c7 = textFormmat("2025/11/18(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"11",
                                "day":"18",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("155960.0",count);
                c3 = textFormmat("86525.0",count);
                c4 = textFormmat("73212.0",count);
                c5 = textFormmat("13313.0",count);
                c6 = textFormmat("18.2",count);
                c7 = textFormmat("2025/10/15(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"10",
                                "day":"15",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("155655.0",count);
                c3 = textFormmat("100937.0",count);
                c4 = textFormmat("89532.0",count);
                c5 = textFormmat("11405.0",count);
                c6 = textFormmat("12.7",count);
                c7 = textFormmat("2025/09/01(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"09",
                                "day":"01",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("155385.0",count);
                c3 = textFormmat("105011.0",count);
                c4 = textFormmat("95951.0",count);
                c5 = textFormmat("9060.0",count);
                c6 = textFormmat("9.4",count);
                c7 = textFormmat("2025/08/25(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"08",
                                "day":"25",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("154955.0",count);
                c3 = textFormmat("105151.0",count);
                c4 = textFormmat("95675.0",count);
                c5 = textFormmat("9476.0",count);
                c6 = textFormmat("9.9",count);
                c7 = textFormmat("2025/07/08(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"07",
                                "day":"08",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("154917.0",count);
                c3 = textFormmat("99615.0",count);
                c4 = textFormmat("85599.0",count);
                c5 = textFormmat("13216.0",count);
                c6 = textFormmat("15.4",count);
                c7 = textFormmat("2025/06/30(19:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"06",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("154557.0",count);
                c3 = textFormmat("82818.0",count);
                c4 = textFormmat("74239.0",count);
                c5 = textFormmat("8579.0",count);
                c6 = textFormmat("11.6",count);
                c7 = textFormmat("2025/05/21(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"05",
                                "day":"21",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("153931.0",count);
                c3 = textFormmat("86002.0",count);
                c4 = textFormmat("71667.0",count);
                c5 = textFormmat("14335.0",count);
                c6 = textFormmat("20",count);
                c7 = textFormmat("2025/04/14(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"04",
                                "day":"14",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("153477.0",count);
                c3 = textFormmat("99269.0",count);
                c4 = textFormmat("85127.0",count);
                c5 = textFormmat("14142.0",count);
                c6 = textFormmat("16.6",count);
                c7 = textFormmat("2025/03/04(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"03",
                                "day":"04",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("153417.0",count);
                c3 = textFormmat("108809.0",count);
                c4 = textFormmat("90638.0",count);
                c5 = textFormmat("18171.0",count);
                c6 = textFormmat("20",count);
                c7 = textFormmat("2025/02/07(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"02",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("152827.0",count);
                c3 = textFormmat("111539.0",count);
                c4 = textFormmat("90705.0",count);
                c5 = textFormmat("20834.0",count);
                c6 = textFormmat("23",count);
                c7 = textFormmat("2025/01/09(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2025",
                                "month":"01",
                                "day":"09",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("153047.0",count);
                c3 = textFormmat("106287.0",count);
                c4 = textFormmat("83314.0",count);
                c5 = textFormmat("22973.0",count);
                c6 = textFormmat("27.6",count);
                c7 = textFormmat("2024/12/19(09:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"12",
                                "day":"19",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("151707.0",count);
                c3 = textFormmat("97490.0",count);
                c4 = textFormmat("79064.0",count);
                c5 = textFormmat("18426.0",count);
                c6 = textFormmat("23.3",count);
                c7 = textFormmat("2024/11/27(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"11",
                                "day":"27",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("150118.0",count);
                c3 = textFormmat("85391.0",count);
                c4 = textFormmat("72272.0",count);
                c5 = textFormmat("13119.0",count);
                c6 = textFormmat("18.2",count);
                c7 = textFormmat("2024/10/15(19:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"10",
                                "day":"15",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("149625.0",count);
                c3 = textFormmat("101806.0",count);
                c4 = textFormmat("93246.0",count);
                c5 = textFormmat("8560.0",count);
                c6 = textFormmat("9.2",count);
                c7 = textFormmat("2024/09/11(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"09",
                                "day":"11",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("149179.0",count);
                c3 = textFormmat("105360.0",count);
                c4 = textFormmat("97115.0",count);
                c5 = textFormmat("8245.0",count);
                c6 = textFormmat("8.5",count);
                c7 = textFormmat("2024/08/20(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"08",
                                "day":"20",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("148932.0",count);
                c3 = textFormmat("99162.0",count);
                c4 = textFormmat("90254.0",count);
                c5 = textFormmat("8908.0",count);
                c6 = textFormmat("9.9",count);
                c7 = textFormmat("2024/07/25(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"07",
                                "day":"25",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("148089.0",count);
                c3 = textFormmat("89170.0",count);
                c4 = textFormmat("80077.0",count);
                c5 = textFormmat("9093.0",count);
                c6 = textFormmat("11.4",count);
                c7 = textFormmat("2024/06/19(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"06",
                                "day":"19",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("147999.0",count);
                c3 = textFormmat("82585.0",count);
                c4 = textFormmat("69899.0",count);
                c5 = textFormmat("12686.0",count);
                c6 = textFormmat("18.1",count);
                c7 = textFormmat("2024/05/23(20:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"05",
                                "day":"23",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("145191.0",count);
                c3 = textFormmat("83155.0",count);
                c4 = textFormmat("71199.0",count);
                c5 = textFormmat("11956.0",count);
                c6 = textFormmat("16.8",count);
                c7 = textFormmat("2024/04/03(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"04",
                                "day":"03",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("145013.0",count);
                c3 = textFormmat("95500.0",count);
                c4 = textFormmat("81691.0",count);
                c5 = textFormmat("13809.0",count);
                c6 = textFormmat("16.9",count);
                c7 = textFormmat("2024/03/05(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"03",
                                "day":"05",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("144983.0",count);
                c3 = textFormmat("101921.0",count);
                c4 = textFormmat("84093.0",count);
                c5 = textFormmat("17828.0",count);
                c6 = textFormmat("21.2",count);
                c7 = textFormmat("2024/02/22(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"02",
                                "day":"22",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("144640.0",count);
                c3 = textFormmat("104997.0",count);
                c4 = textFormmat("89231.0",count);
                c5 = textFormmat("15766.0",count);
                c6 = textFormmat("17.7",count);
                c7 = textFormmat("2024/01/23(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2024",
                                "month":"01",
                                "day":"23",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("144381.0",count);
                c3 = textFormmat("105213.0",count);
                c4 = textFormmat("91556.0",count);
                c5 = textFormmat("13657.0",count);
                c6 = textFormmat("14.9",count);
                c7 = textFormmat("2023/12/21(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"12",
                                "day":"21",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("144063.0",count);
                c3 = textFormmat("96366.0",count);
                c4 = textFormmat("82716.0",count);
                c5 = textFormmat("13650.0",count);
                c6 = textFormmat("16.5",count);
                c7 = textFormmat("2023/11/30(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"11",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("143712.0",count);
                c3 = textFormmat("85738.0",count);
                c4 = textFormmat("70985.0",count);
                c5 = textFormmat("14753.0",count);
                c6 = textFormmat("20.8",count);
                c7 = textFormmat("2023/10/19(19:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"10",
                                "day":"19",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("143401.0",count);
                c3 = textFormmat("99391.0",count);
                c4 = textFormmat("85682.0",count);
                c5 = textFormmat("13709.0",count);
                c6 = textFormmat("16",count);
                c7 = textFormmat("2023/09/05(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"09",
                                "day":"05",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("142567.0",count);
                c3 = textFormmat("104297.0",count);
                c4 = textFormmat("93615.0",count);
                c5 = textFormmat("10682.0",count);
                c6 = textFormmat("11.4",count);
                c7 = textFormmat("2023/08/07(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"08",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("142201.0",count);
                c3 = textFormmat("102234.0",count);
                c4 = textFormmat("87033.0",count);
                c5 = textFormmat("15201.0",count);
                c6 = textFormmat("17.5",count);
                c7 = textFormmat("2023/07/27(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"07",
                                "day":"27",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("140955.0",count);
                c3 = textFormmat("93562.0",count);
                c4 = textFormmat("82196.0",count);
                c5 = textFormmat("11366.0",count);
                c6 = textFormmat("13.8",count);
                c7 = textFormmat("2023/06/29(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"06",
                                "day":"29",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("140240.0",count);
                c3 = textFormmat("82332.0",count);
                c4 = textFormmat("72772.0",count);
                c5 = textFormmat("9560.0",count);
                c6 = textFormmat("13.1",count);
                c7 = textFormmat("2023/05/30(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"05",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("139079.0",count);
                c3 = textFormmat("84179.0",count);
                c4 = textFormmat("72292.0",count);
                c5 = textFormmat("11887.0",count);
                c6 = textFormmat("16.4",count);
                c7 = textFormmat("2023/04/05(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"04",
                                "day":"05",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("139079.0",count);
                c3 = textFormmat("84179.0",count);
                c4 = textFormmat("72292.0",count);
                c5 = textFormmat("11887.0",count);
                c6 = textFormmat("16.4",count);
                c7 = textFormmat("2023/04/05(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"04",
                                "day":"05",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("138860.0",count);
                c3 = textFormmat("97952.0",count);
                c4 = textFormmat("76510.0",count);
                c5 = textFormmat("21442.0",count);
                c6 = textFormmat("28",count);
                c7 = textFormmat("2023/03/02(19:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"03",
                                "day":"02",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("138647.0",count);
                c3 = textFormmat("101103.0",count);
                c4 = textFormmat("84290.0",count);
                c5 = textFormmat("16813.0",count);
                c6 = textFormmat("19.9",count);
                c7 = textFormmat("2023/02/03(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"02",
                                "day":"03",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("138399.0",count);
                c3 = textFormmat("103138.0",count);
                c4 = textFormmat("92613.0",count);
                c5 = textFormmat("10525.0",count);
                c6 = textFormmat("11.4",count);
                c7 = textFormmat("2023/01/26(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2023",
                                "month":"01",
                                "day":"26",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("137938.0",count);
                c3 = textFormmat("105628.0",count);
                c4 = textFormmat("94509.0",count);
                c5 = textFormmat("11119.0",count);
                c6 = textFormmat("11.8",count);
                c7 = textFormmat("2022/12/23(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"12",
                                "day":"23",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("136268.0",count);
                c3 = textFormmat("92682.0",count);
                c4 = textFormmat("82117.0",count);
                c5 = textFormmat("10565.0",count);
                c6 = textFormmat("12.9",count);
                c7 = textFormmat("2022/11/30(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"11",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("134768.0",count);
                c3 = textFormmat("86098.0",count);
                c4 = textFormmat("72544.0",count);
                c5 = textFormmat("13554.0",count);
                c6 = textFormmat("18.7",count);
                c7 = textFormmat("2022/10/04(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"10",
                                "day":"04",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("134719.0",count);
                c3 = textFormmat("91923.0",count);
                c4 = textFormmat("82122.0",count);
                c5 = textFormmat("9801.0",count);
                c6 = textFormmat("11.9",count);
                c7 = textFormmat("2022/09/16(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"09",
                                "day":"16",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("134417.0",count);
                c3 = textFormmat("100691.0",count);
                c4 = textFormmat("89263.0",count);
                c5 = textFormmat("11428.0",count);
                c6 = textFormmat("12.8",count);
                c7 = textFormmat("2022/08/08(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"08",
                                "day":"08",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("134239.0",count);
                c3 = textFormmat("99716.0",count);
                c4 = textFormmat("92990.0",count);
                c5 = textFormmat("6726.0",count);
                c6 = textFormmat("7.2",count);
                c7 = textFormmat("2022/07/07(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"07",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("134092.0",count);
                c3 = textFormmat("94364.0",count);
                c4 = textFormmat("84739.0",count);
                c5 = textFormmat("9625.0",count);
                c6 = textFormmat("11.4",count);
                c7 = textFormmat("2022/06/27(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"06",
                                "day":"27",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("133917.0",count);
                c3 = textFormmat("84474.0",count);
                c4 = textFormmat("73134.0",count);
                c5 = textFormmat("11340.0",count);
                c6 = textFormmat("15.5",count);
                c7 = textFormmat("2022/05/30(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"05",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("133680.0",count);
                c3 = textFormmat("84457.0",count);
                c4 = textFormmat("71879.0",count);
                c5 = textFormmat("12578.0",count);
                c6 = textFormmat("17.5",count);
                c7 = textFormmat("2022/04/13(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"04",
                                "day":"13",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("133561.0",count);
                c3 = textFormmat("89033.0",count);
                c4 = textFormmat("78233.0",count);
                c5 = textFormmat("10800.0",count);
                c6 = textFormmat("13.8",count);
                c7 = textFormmat("2022/03/23(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"03",
                                "day":"23",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("133069.0",count);
                c3 = textFormmat("100103.0",count);
                c4 = textFormmat("87351.0",count);
                c5 = textFormmat("12752.0",count);
                c6 = textFormmat("14.6",count);
                c7 = textFormmat("2022/02/07(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"02",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("134020.0",count);
                c3 = textFormmat("107631.0",count);
                c4 = textFormmat("89397.0",count);
                c5 = textFormmat("18234.0",count);
                c6 = textFormmat("20.4",count);
                c7 = textFormmat("2022/01/05(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2022",
                                "month":"01",
                                "day":"05",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("134158.0",count);
                c3 = textFormmat("103554.0",count);
                c4 = textFormmat("90708.0",count);
                c5 = textFormmat("12846.0",count);
                c6 = textFormmat("14.2",count);
                c7 = textFormmat("2021/12/27(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"12",
                                "day":"27",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("133564.0",count);
                c3 = textFormmat("95394.0",count);
                c4 = textFormmat("80362.0",count);
                c5 = textFormmat("15032.0",count);
                c6 = textFormmat("18.7",count);
                c7 = textFormmat("2021/11/30(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"11",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("131895.0",count);
                c3 = textFormmat("82449.0",count);
                c4 = textFormmat("75698.0",count);
                c5 = textFormmat("6751.0",count);
                c6 = textFormmat("8.9",count);
                c7 = textFormmat("2021/10/05(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"10",
                                "day":"05",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("131895.0",count);
                c3 = textFormmat("90532.0",count);
                c4 = textFormmat("77820.0",count);
                c5 = textFormmat("12712.0",count);
                c6 = textFormmat("16.3",count);
                c7 = textFormmat("2021/09/13(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"09",
                                "day":"13",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("131330.0",count);
                c3 = textFormmat("99241.0",count);
                c4 = textFormmat("80362.0",count);
                c5 = textFormmat("18879.0",count);
                c6 = textFormmat("23.5",count);
                c7 = textFormmat("2021/08/18(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"08",
                                "day":"18",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("131330.0",count);
                c3 = textFormmat("98952.0",count);
                c4 = textFormmat("86355.0",count);
                c5 = textFormmat("12597.0",count);
                c6 = textFormmat("14.6",count);
                c7 = textFormmat("2021/08/12(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"08",
                                "day":"12",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("131330.0",count);
                c3 = textFormmat("100739.0",count);
                c4 = textFormmat("91141.0",count);
                c5 = textFormmat("9598.0",count);
                c6 = textFormmat("10.5",count);
                c7 = textFormmat("2021/07/27(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"07",
                                "day":"27",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("129623.0",count);
                c3 = textFormmat("87573.0",count);
                c4 = textFormmat("75854.0",count);
                c5 = textFormmat("11719.0",count);
                c6 = textFormmat("15.4",count);
                c7 = textFormmat("2021/06/29(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"06",
                                "day":"29",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("129540",count);
                c3 = textFormmat("80387.0",count);
                c4 = textFormmat("69140",count);
                c5 = textFormmat("11247.0",count);
                c6 = textFormmat("16.3",count);
                c7 = textFormmat("2021/05/20(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"05",
                                "day":"20",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("129361.0",count);
                c3 = textFormmat("77695.0",count);
                c4 = textFormmat("68966.0",count);
                c5 = textFormmat("8729.0",count);
                c6 = textFormmat("12.7",count);
                c7 = textFormmat("2021/04/12(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"04",
                                "day":"12",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("128820",count);
                c3 = textFormmat("92526.0",count);
                c4 = textFormmat("77208.0",count);
                c5 = textFormmat("15318.0",count);
                c6 = textFormmat("19.8",count);
                c7 = textFormmat("2021/03/02(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"03",
                                "day":"02",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("128820",count);
                c3 = textFormmat("95992.0",count);
                c4 = textFormmat("84749.0",count);
                c5 = textFormmat("11243.0",count);
                c6 = textFormmat("13.3",count);
                c7 = textFormmat("2021/02/17(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"02",
                                "day":"17",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("128209.0",count);
                c3 = textFormmat("99189.0",count);
                c4 = textFormmat("90564.0",count);
                c5 = textFormmat("8625.0",count);
                c6 = textFormmat("9.5",count);
                c7 = textFormmat("2021/01/11(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2021",
                                "month":"01",
                                "day":"11",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("128609.0",count);
                c3 = textFormmat("98795.0",count);
                c4 = textFormmat("85132.0",count);
                c5 = textFormmat("13663.0",count);
                c6 = textFormmat("16",count);
                c7 = textFormmat("2020/12/16(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"12",
                                "day":"16",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("128609.0",count);
                c3 = textFormmat("96827.0",count);
                c4 = textFormmat("77074.0",count);
                c5 = textFormmat("19753.0",count);
                c6 = textFormmat("25.6",count);
                c7 = textFormmat("2020/11/30(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"11",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("127762.0",count);
                c3 = textFormmat("85999.0",count);
                c4 = textFormmat("68454.0",count);
                c5 = textFormmat("17545.0",count);
                c6 = textFormmat("25.6",count);
                c7 = textFormmat("2020/10/21(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"10",
                                "day":"21",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("128078.0",count);
                c3 = textFormmat("96797.0",count);
                c4 = textFormmat("82036.0",count);
                c5 = textFormmat("14761.0",count);
                c6 = textFormmat("18",count);
                c7 = textFormmat("2020/09/01(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"09",
                                "day":"01",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("127819.0",count);
                c3 = textFormmat("97951.0",count);
                c4 = textFormmat("89091.0",count);
                c5 = textFormmat("8860",count);
                c6 = textFormmat("9.9",count);
                c7 = textFormmat("2020/08/26(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"08",
                                "day":"26",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("127338.0",count);
                c3 = textFormmat("97338.0",count);
                c4 = textFormmat("75675.0",count);
                c5 = textFormmat("21663.0",count);
                c6 = textFormmat("28.6",count);
                c7 = textFormmat("2020/07/09(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"07",
                                "day":"09",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("126798.0",count);
                c3 = textFormmat("91874.0",count);
                c4 = textFormmat("75247.0",count);
                c5 = textFormmat("16627.0",count);
                c6 = textFormmat("22.1",count);
                c7 = textFormmat("2020/06/10(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"06",
                                "day":"10",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("126580",count);
                c3 = textFormmat("88211.0",count);
                c4 = textFormmat("65700",count);
                c5 = textFormmat("22511.0",count);
                c6 = textFormmat("34.3",count);
                c7 = textFormmat("2020/05/18(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"05",
                                "day":"18",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("126262.0",count);
                c3 = textFormmat("83368.0",count);
                c4 = textFormmat("66699.0",count);
                c5 = textFormmat("16669.0",count);
                c6 = textFormmat("25",count);
                c7 = textFormmat("2020/04/17(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"04",
                                "day":"17",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("125887.0",count);
                c3 = textFormmat("90877.0",count);
                c4 = textFormmat("73329.0",count);
                c5 = textFormmat("17548.0",count);
                c6 = textFormmat("23.9",count);
                c7 = textFormmat("2020/03/10(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"03",
                                "day":"10",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("125358.0",count);
                c3 = textFormmat("97069.0",count);
                c4 = textFormmat("81493.0",count);
                c5 = textFormmat("15576.0",count);
                c6 = textFormmat("19.1",count);
                c7 = textFormmat("2020/02/06(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"02",
                                "day":"06",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("125358.0",count);
                c3 = textFormmat("94735.0",count);
                c4 = textFormmat("82352.0",count);
                c5 = textFormmat("12383.0",count);
                c6 = textFormmat("15",count);
                c7 = textFormmat("2020/01/16(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2020",
                                "month":"01",
                                "day":"16",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("124833.0",count);
                c3 = textFormmat("92934.0",count);
                c4 = textFormmat("81787.0",count);
                c5 = textFormmat("11147.0",count);
                c6 = textFormmat("13.6",count);
                c7 = textFormmat("2019/12/06(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"12",
                                "day":"06",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("124392.0",count);
                c3 = textFormmat("87377.0",count);
                c4 = textFormmat("74393.0",count);
                c5 = textFormmat("12984.0",count);
                c6 = textFormmat("17.5",count);
                c7 = textFormmat("2019/11/19(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"11",
                                "day":"19",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("122490",count);
                c3 = textFormmat("83471.0",count);
                c4 = textFormmat("73550",count);
                c5 = textFormmat("9921.0",count);
                c6 = textFormmat("13.5",count);
                c7 = textFormmat("2019/10/01(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"10",
                                "day":"01",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("122973.0",count);
                c3 = textFormmat("96032.0",count);
                c4 = textFormmat("79916.0",count);
                c5 = textFormmat("16116.0",count);
                c6 = textFormmat("20.2",count);
                c7 = textFormmat("2019/09/06(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"09",
                                "day":"06",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("121147.0",count);
                c3 = textFormmat("96389.0",count);
                c4 = textFormmat("90314.0",count);
                c5 = textFormmat("6075.0",count);
                c6 = textFormmat("6.7",count);
                c7 = textFormmat("2019/08/13(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"08",
                                "day":"13",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("121147.0",count);
                c3 = textFormmat("95593.0",count);
                c4 = textFormmat("84164.0",count);
                c5 = textFormmat("11429.0",count);
                c6 = textFormmat("13.6",count);
                c7 = textFormmat("2019/07/23(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"07",
                                "day":"23",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("120471.0",count);
                c3 = textFormmat("90520",count);
                c4 = textFormmat("75206.0",count);
                c5 = textFormmat("15314.0",count);
                c6 = textFormmat("20.4",count);
                c7 = textFormmat("2019/06/27(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"06",
                                "day":"27",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("119826.0",count);
                c3 = textFormmat("86270",count);
                c4 = textFormmat("70964.0",count);
                c5 = textFormmat("15306.0",count);
                c6 = textFormmat("21.6",count);
                c7 = textFormmat("2019/05/24(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"05",
                                "day":"24",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("119826.0",count);
                c3 = textFormmat("86455.0",count);
                c4 = textFormmat("73010",count);
                c5 = textFormmat("13445.0",count);
                c6 = textFormmat("18.4",count);
                c7 = textFormmat("2019/04/10(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"04",
                                "day":"10",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("119350",count);
                c3 = textFormmat("90422.0",count);
                c4 = textFormmat("74136.0",count);
                c5 = textFormmat("16286.0",count);
                c6 = textFormmat("22",count);
                c7 = textFormmat("2019/03/13(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"03",
                                "day":"13",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("119350",count);
                c3 = textFormmat("98738.0",count);
                c4 = textFormmat("82869.0",count);
                c5 = textFormmat("15869.0",count);
                c6 = textFormmat("19.1",count);
                c7 = textFormmat("2019/02/19(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"02",
                                "day":"19",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("119092.0",count);
                c3 = textFormmat("100827.0",count);
                c4 = textFormmat("85392.0",count);
                c5 = textFormmat("15435.0",count);
                c6 = textFormmat("18.1",count);
                c7 = textFormmat("2019/01/09(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2019",
                                "month":"01",
                                "day":"09",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("119092.0",count);
                c3 = textFormmat("100841.0",count);
                c4 = textFormmat("86083.0",count);
                c5 = textFormmat("14758.0",count);
                c6 = textFormmat("17.1",count);
                c7 = textFormmat("2018/12/28(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"12",
                                "day":"28",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("118266.0",count);
                c3 = textFormmat("93480",count);
                c4 = textFormmat("75914.0",count);
                c5 = textFormmat("17566.0",count);
                c6 = textFormmat("23.1",count);
                c7 = textFormmat("2018/11/29(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"11",
                                "day":"29",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("117996.0",count);
                c3 = textFormmat("84796.0",count);
                c4 = textFormmat("70329.0",count);
                c5 = textFormmat("14467.0",count);
                c6 = textFormmat("20.6",count);
                c7 = textFormmat("2018/10/30(19:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"10",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("117996.0",count);
                c3 = textFormmat("94223.0",count);
                c4 = textFormmat("78405.0",count);
                c5 = textFormmat("15818.0",count);
                c6 = textFormmat("20.2",count);
                c7 = textFormmat("2018/09/03(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"09",
                                "day":"03",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("117840",count);
                c3 = textFormmat("100250",count);
                c4 = textFormmat("91548.0",count);
                c5 = textFormmat("8702.0",count);
                c6 = textFormmat("9.5",count);
                c7 = textFormmat("2018/08/14(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"08",
                                "day":"14",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("117205.0",count);
                c3 = textFormmat("99570",count);
                c4 = textFormmat("92478.0",count);
                c5 = textFormmat("7092.0",count);
                c6 = textFormmat("7.7",count);
                c7 = textFormmat("2018/07/24(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"07",
                                "day":"24",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("117111.0",count);
                c3 = textFormmat("87256.0",count);
                c4 = textFormmat("76606.0",count);
                c5 = textFormmat("10650",count);
                c6 = textFormmat("13.9",count);
                c7 = textFormmat("2018/06/25(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"06",
                                "day":"25",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("117589.0",count);
                c3 = textFormmat("82000",count);
                c4 = textFormmat("73173.0",count);
                c5 = textFormmat("8827.0",count);
                c6 = textFormmat("12.1",count);
                c7 = textFormmat("2018/05/17(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"05",
                                "day":"17",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("116747.0",count);
                c3 = textFormmat("84634.0",count);
                c4 = textFormmat("70520",count);
                c5 = textFormmat("14114.0",count);
                c6 = textFormmat("20",count);
                c7 = textFormmat("2018/04/05(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"04",
                                "day":"05",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("116747.0",count);
                c3 = textFormmat("90761.0",count);
                c4 = textFormmat("78855.0",count);
                c5 = textFormmat("11906.0",count);
                c6 = textFormmat("15.1",count);
                c7 = textFormmat("2018/03/08(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"03",
                                "day":"08",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("116428.0",count);
                c3 = textFormmat("101148.0",count);
                c4 = textFormmat("88238.0",count);
                c5 = textFormmat("12910",count);
                c6 = textFormmat("14.6",count);
                c7 = textFormmat("2018/02/06(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"02",
                                "day":"06",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("116428.0",count);
                c3 = textFormmat("99147.0",count);
                c4 = textFormmat("87247.0",count);
                c5 = textFormmat("11900",count);
                c6 = textFormmat("13.6",count);
                c7 = textFormmat("2018/01/25(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2018",
                                "month":"01",
                                "day":"25",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("116657.0",count);
                c3 = textFormmat("96095.0",count);
                c4 = textFormmat("85133.0",count);
                c5 = textFormmat("10962.0",count);
                c6 = textFormmat("12.9",count);
                c7 = textFormmat("2017/12/12(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"12",
                                "day":"12",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("116278.0",count);
                c3 = textFormmat("92807.0",count);
                c4 = textFormmat("77605.0",count);
                c5 = textFormmat("15202.0",count);
                c6 = textFormmat("19.6",count);
                c7 = textFormmat("2017/11/24(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"11",
                                "day":"24",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("115249.0",count);
                c3 = textFormmat("88736.0",count);
                c4 = textFormmat("70113.0",count);
                c5 = textFormmat("18623.0",count);
                c6 = textFormmat("26.6",count);
                c7 = textFormmat("2017/10/10(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"10",
                                "day":"10",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("114206.0",count);
                c3 = textFormmat("92665.0",count);
                c4 = textFormmat("73115.0",count);
                c5 = textFormmat("19550",count);
                c6 = textFormmat("26.7",count);
                c7 = textFormmat("2017/09/19(16:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"09",
                                "day":"19",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("113351.0",count);
                c3 = textFormmat("95125.0",count);
                c4 = textFormmat("84489.0",count);
                c5 = textFormmat("10636.0",count);
                c6 = textFormmat("12.6",count);
                c7 = textFormmat("2017/08/07(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"08",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("113351.0",count);
                c3 = textFormmat("94987.0",count);
                c4 = textFormmat("84586.0",count);
                c5 = textFormmat("10401.0",count);
                c6 = textFormmat("12.3",count);
                c7 = textFormmat("2017/07/21(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"07",
                                "day":"21",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("113705.0",count);
                c3 = textFormmat("87888.0",count);
                c4 = textFormmat("75413.0",count);
                c5 = textFormmat("12475.0",count);
                c6 = textFormmat("16.5",count);
                c7 = textFormmat("2017/06/30(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"06",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("111292.0",count);
                c3 = textFormmat("82497.0",count);
                c4 = textFormmat("69887.0",count);
                c5 = textFormmat("12610",count);
                c6 = textFormmat("18",count);
                c7 = textFormmat("2017/05/30(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"05",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("110502.0",count);
                c3 = textFormmat("87916.0",count);
                c4 = textFormmat("69129.0",count);
                c5 = textFormmat("18787.0",count);
                c6 = textFormmat("27.2",count);
                c7 = textFormmat("2017/04/06(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"04",
                                "day":"06",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("107098.0",count);
                c3 = textFormmat("93840",count);
                c4 = textFormmat("76756.0",count);
                c5 = textFormmat("17084.0",count);
                c6 = textFormmat("22.3",count);
                c7 = textFormmat("2017/03/07(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"03",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("107098.0",count);
                c3 = textFormmat("95916.0",count);
                c4 = textFormmat("81268.0",count);
                c5 = textFormmat("14648.0",count);
                c6 = textFormmat("18",count);
                c7 = textFormmat("2017/02/10(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"02",
                                "day":"10",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("106238.0",count);
                c3 = textFormmat("95443.0",count);
                c4 = textFormmat("83657.0",count);
                c5 = textFormmat("11786.0",count);
                c6 = textFormmat("14.1",count);
                c7 = textFormmat("2017/01/23(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2017",
                                "month":"01",
                                "day":"23",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("104466.0",count);
                c3 = textFormmat("93031.0",count);
                c4 = textFormmat("79966.0",count);
                c5 = textFormmat("13065.0",count);
                c6 = textFormmat("16.3",count);
                c7 = textFormmat("2016/12/16(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"12",
                                "day":"16",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("103299.0",count);
                c3 = textFormmat("85024.0",count);
                c4 = textFormmat("74959.0",count);
                c5 = textFormmat("10065.0",count);
                c6 = textFormmat("13.4",count);
                c7 = textFormmat("2016/11/30(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"11",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("103068.0",count);
                c3 = textFormmat("78563.0",count);
                c4 = textFormmat("68282.0",count);
                c5 = textFormmat("10281.0",count);
                c6 = textFormmat("15.1",count);
                c7 = textFormmat("2016/10/31(18:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"10",
                                "day":"31",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("101025.0",count);
                c3 = textFormmat("85109.0",count);
                c4 = textFormmat("75299.0",count);
                c5 = textFormmat("9810",count);
                c6 = textFormmat("13",count);
                c7 = textFormmat("2016/09/05(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"09",
                                "day":"05",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("100180",count);
                c3 = textFormmat("92395.0",count);
                c4 = textFormmat("85183.0",count);
                c5 = textFormmat("7212.0",count);
                c6 = textFormmat("8.5",count);
                c7 = textFormmat("2016/08/12(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"08",
                                "day":"12",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("100180",count);
                c3 = textFormmat("88923.0",count);
                c4 = textFormmat("81110",count);
                c5 = textFormmat("7813.0",count);
                c6 = textFormmat("9.6",count);
                c7 = textFormmat("2016/07/26(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"07",
                                "day":"26",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("98948.0",count);
                c3 = textFormmat("80956.0",count);
                c4 = textFormmat("72816.0",count);
                c5 = textFormmat("8140",count);
                c6 = textFormmat("11.2",count);
                c7 = textFormmat("2016/06/23(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"06",
                                "day":"23",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("98948.0",count);
                c3 = textFormmat("80087.0",count);
                c4 = textFormmat("68011.0",count);
                c5 = textFormmat("12076.0",count);
                c6 = textFormmat("17.8",count);
                c7 = textFormmat("2016/05/31(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"05",
                                "day":"31",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("98796.0",count);
                c3 = textFormmat("86048.0",count);
                c4 = textFormmat("65765.0",count);
                c5 = textFormmat("20283.0",count);
                c6 = textFormmat("30.8",count);
                c7 = textFormmat("2016/04/07(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"04",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("98796.0",count);
                c3 = textFormmat("85097.0",count);
                c4 = textFormmat("75697.0",count);
                c5 = textFormmat("9400",count);
                c6 = textFormmat("12.4",count);
                c7 = textFormmat("2016/03/02(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"03",
                                "day":"02",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("98190",count);
                c3 = textFormmat("93809.0",count);
                c4 = textFormmat("80020",count);
                c5 = textFormmat("13789.0",count);
                c6 = textFormmat("17.2",count);
                c7 = textFormmat("2016/02/02(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"02",
                                "day":"02",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("98190",count);
                c3 = textFormmat("94793.0",count);
                c4 = textFormmat("82972.0",count);
                c5 = textFormmat("11821.0",count);
                c6 = textFormmat("14.2",count);
                c7 = textFormmat("2016/01/21(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2016",
                                "month":"01",
                                "day":"21",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("97649.0",count);
                c3 = textFormmat("89191.0",count);
                c4 = textFormmat("77282.0",count);
                c5 = textFormmat("11909.0",count);
                c6 = textFormmat("15.4",count);
                c7 = textFormmat("2015/12/18(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"12",
                                "day":"18",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("97649.0",count);
                c3 = textFormmat("87180",count);
                c4 = textFormmat("73965.0",count);
                c5 = textFormmat("13215.0",count);
                c6 = textFormmat("17.9",count);
                c7 = textFormmat("2015/11/27(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"11",
                                "day":"27",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("96828.0",count);
                c3 = textFormmat("80604.0",count);
                c4 = textFormmat("65322.0",count);
                c5 = textFormmat("15282.0",count);
                c6 = textFormmat("23.4",count);
                c7 = textFormmat("2015/10/01(17:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"10",
                                "day":"01",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("96828.0",count);
                c3 = textFormmat("85102.0",count);
                c4 = textFormmat("72991.0",count);
                c5 = textFormmat("12111.0",count);
                c6 = textFormmat("16.6",count);
                c7 = textFormmat("2015/09/01(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"09",
                                "day":"01",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("96828.0",count);
                c3 = textFormmat("89595.0",count);
                c4 = textFormmat("76916.0",count);
                c5 = textFormmat("12679.0",count);
                c6 = textFormmat("16.5",count);
                c7 = textFormmat("2015/08/07(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"08",
                                "day":"07",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("96828.0",count);
                c3 = textFormmat("89654.0",count);
                c4 = textFormmat("76696.0",count);
                c5 = textFormmat("12958.0",count);
                c6 = textFormmat("16.9",count);
                c7 = textFormmat("2015/07/30(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"07",
                                "day":"30",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("95681.0",count);
                c3 = textFormmat("84047.0",count);
                c4 = textFormmat("69882.0",count);
                c5 = textFormmat("14165.0",count);
                c6 = textFormmat("20.3",count);
                c7 = textFormmat("2015/06/24(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"06",
                                "day":"24",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("95681.0",count);
                c3 = textFormmat("76874.0",count);
                c4 = textFormmat("67107.0",count);
                c5 = textFormmat("9767.0",count);
                c6 = textFormmat("14.6",count);
                c7 = textFormmat("2015/05/29(15:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"05",
                                "day":"29",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("95365.0",count);
                c3 = textFormmat("79635.0",count);
                c4 = textFormmat("66582.0",count);
                c5 = textFormmat("13053.0",count);
                c6 = textFormmat("19.6",count);
                c7 = textFormmat("2015/04/08(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"04",
                                "day":"08",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("94102.0",count);
                c3 = textFormmat("84910",count);
                c4 = textFormmat("75375.0",count);
                c5 = textFormmat("9535.0",count);
                c6 = textFormmat("12.7",count);
                c7 = textFormmat("2015/03/10(01:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"03",
                                "day":"10",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("94102.0",count);
                c3 = textFormmat("87926.0",count);
                c4 = textFormmat("78790",count);
                c5 = textFormmat("9136.0",count);
                c6 = textFormmat("11.6",count);
                c7 = textFormmat("2015/02/09(11:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"02",
                                "day":"09",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            
                c1 = textFormmat("92825.0",count);
                c3 = textFormmat("90501.0",count);
                c4 = textFormmat("77796.0",count);
                c5 = textFormmat("12705.0",count);
                c6 = textFormmat("16.3",count);
                c7 = textFormmat("2015/01/08(10:00)",count);
                c8 = ("0" == 0) ? "-" : Math.round("0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("" == "") ? "-" : textFormmat("",count);

                gridData.push({"year":"2015",
                                "month":"01",
                                "day":"08",
                                "c1":c1,
                                "c3":c3,
                                "c4":c4,
                                "c8":c8,
                                "c9":c9,
                                "c5":c5,
                                "c6":c6,
                                "c7":c7
                                });
            

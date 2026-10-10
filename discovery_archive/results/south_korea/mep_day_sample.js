









            var c1,c2,c3,c4,c5,c6,c8,c9;
            var count = $("#selCount option:selected").val();
            var point = "";
            if(count == 0) point = "";
            else if(count == 1) point = ".0";
            else if(count == 2) point = ".00";
            else if(count == 3) point = ".000";
            gridData = [];

            
                c1 = textFormmat("161493.0",count);
                c3 = textFormmat("88443.0",count);
                c4 = textFormmat("65158.0",count);
                c5 = textFormmat("23285.0",count);
                c6 = textFormmat("35.7",count);
                c7 = textFormmat("2026/10/09(19:00)",count);
                c8 = ("47116.0" == 0) ? "-" : Math.round("47116.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/10/09(13:00)" == "") ? "-" : textFormmat("2026/10/09(13:00)",count);

                gridData.push({"year":"2026",
                                "month":"10",
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
            
                c1 = textFormmat("161493.0",count);
                c3 = textFormmat("94930.0",count);
                c4 = textFormmat("71154.0",count);
                c5 = textFormmat("23776.0",count);
                c6 = textFormmat("33.4",count);
                c7 = textFormmat("2026/10/08(19:00)",count);
                c8 = ("51395.0" == 0) ? "-" : Math.round("51395.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/10/08(13:00)" == "") ? "-" : textFormmat("2026/10/08(13:00)",count);

                gridData.push({"year":"2026",
                                "month":"10",
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
            
                c1 = textFormmat("161493.0",count);
                c3 = textFormmat("96456.0",count);
                c4 = textFormmat("71028.0",count);
                c5 = textFormmat("25428.0",count);
                c6 = textFormmat("35.8",count);
                c7 = textFormmat("2026/10/07(19:00)",count);
                c8 = ("51346.0" == 0) ? "-" : Math.round("51346.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/10/07(13:00)" == "") ? "-" : textFormmat("2026/10/07(13:00)",count);

                gridData.push({"year":"2026",
                                "month":"10",
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
            
                c1 = textFormmat("161478.0",count);
                c3 = textFormmat("96621.0",count);
                c4 = textFormmat("70838.0",count);
                c5 = textFormmat("25783.0",count);
                c6 = textFormmat("36.4",count);
                c7 = textFormmat("2026/10/06(19:00)",count);
                c8 = ("48821.0" == 0) ? "-" : Math.round("48821.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/10/06(13:00)" == "") ? "-" : textFormmat("2026/10/06(13:00)",count);

                gridData.push({"year":"2026",
                                "month":"10",
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
            
                c1 = textFormmat("161478.0",count);
                c3 = textFormmat("93500.0",count);
                c4 = textFormmat("64556.0",count);
                c5 = textFormmat("28944.0",count);
                c6 = textFormmat("44.8",count);
                c7 = textFormmat("2026/10/05(19:00)",count);
                c8 = ("48897.0" == 0) ? "-" : Math.round("48897.0").toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + point;
                c9 = ("2026/10/05(04:00)" == "") ? "-" : textFormmat("2026/10/05(04:00)",count);

                gridData.push({"year":"2026",
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
            
                c1 = textFormmat("161478.0",count);
                c3 = textFormmat("91988.0",count);
                c4 = textFormmat("61713.0",count);
                c5 = textFormmat("30275.0",count);
                c6 = textFormmat("49.1",count);

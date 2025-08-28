/* Formula
=SE(E($C6<=HOJE(); $C6=DATA(2025;H$1;H$5); $E6<=HOJE(); $E6<>"");1;
    SE(E($F6<=HOJE(); $F6=DATA(2025;H$1;H$5)); 3;
        SE(HOJE()<=(DATA(2025;H$1;H$5)-1); "";
            SE(E(G6<>3;$E6+1=DATA(2025;H$1;H$5)); 2;
               SE(G6=1;1;
                  SE(G6=2;2;
                     SE(E(G6=3;);""; "")
                  )
               )
            )
        )
    )
)
*/
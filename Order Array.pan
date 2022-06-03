2D array 
rows are ¶ delimited
columns are ¬ delimited 
 
        y                                    x
  1)	4055-A	  1	  1	organic       	   2.50	   2.50	Rutgers Tomato OG
  2)	4149-A	  1	  1	organic       	   3.00	   3.00	Heirloom Tomato Mix OG
  3)	4250-A	  1	  1	              	   3.25	   3.25	Sun Gold Cherry Tomato
  4)	4920-A	  1	  1	              	   2.50	   2.50	Pacific Beauty Calendula
  5)	5224-B	  1	  1	              	   4.25	   4.25	Brocade Mix Marigold
  6)	5634-A	  1	  1	              	   2.50	   2.50	Kaleidoscope Sweet Pea
  7)	5649-A	  1	  1	              	   2.50	   2.50	Torch Tithonia
  8)	5708-B	  1	  1	organic       	   7.00	   7.00	Benarys Gt Mix Zinnia OG
  9)	5713-A	  1	  1	organic       	   3.25	   3.25	CA Giant Zinnia Mix OG
 10)	5748-A	  1	  1	              	   2.75	   2.75	Jazzy Mix Zinnia

 Iterator is the list of the first part of this array

 what I need at the end: 
The x for where the y is not in the range 1-4699, and 5932-5939

arraysize and arrayrange




          //TEST: Get range 4700-5931, exclude 5932-5939,get above 5940


           10)	5748-A	  1	  1	              	   2.75	   2.75	Jazzy Mix Zinnia^  9)	5713-A	  1	  1	organic       	   3.25	   3.25	CA Giant Zinnia Mix OG^  8)	5708-B	  1	  1	organic       	   7.00	   7.00	Benarys Gt Mix Zinnia OG^  7)	5649-A	  1	  1	              	   2.50	   2.50	Torch Tithonia^  6)	5634-A	  1	  1	              	   2.50	   2.50	Kaleidoscope Sweet Pea^  5)	5224-B	  1	  1	              	   4.25	   4.25	Brocade Mix Marigold^  4)	4920-A	  1	  1	              	   2.50	   2.50	Pacific Beauty Calendula^  3)	4250-A	  1	  1	              	   3.25	   3.25	Sun Gold Cherry Tomato^  2)	4149-A	  1	  1	organic       	   3.00	   3.00	Heirloom Tomato Mix OG^  1)	4055-A	  1	  1	organic       	   2.50	   2.50	Rutgers Tomato OG^

           we made a list delimted by ^ 

           now, we just need to take that list and grab the 

           5748-A^5713-A^5708-B^5649-A^5634-A^5224-B^4920-A^4250-A^4149-A^4055-A^
           5748^5713^5708^5649^5634^5224^4920^4250^4149^4055^
           ^4055^4149^4250^4920^5224^5634^5649^5708^5713^5748

           5748^5713^5708^5649^5634^5224^4920^
           5748	   2.75^5713	   3.25^5708	   7.00^5649	   2.50^5634	   2.50^5224	   4.25^4920	   2.50^
           ^4920	   2.50^5224	   4.25^5634	   2.50^5649	   2.50^5708	   7.00^5713	   3.25^5748	   2.75
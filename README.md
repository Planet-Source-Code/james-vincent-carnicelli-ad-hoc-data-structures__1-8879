<div align="center">

## Ad Hoc Data Structures


</div>

### Description

Simple ways to create data structures on the fly without creating classes, user-defined types, etc., and how such structures can be more flexible, data-driven, and even self-defining.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James Vincent Carnicelli](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-vincent-carnicelli.md)
**Level**          |Intermediate
**User Rating**    |4.8 (76 globes from 16 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-vincent-carnicelli-ad-hoc-data-structures__1-8879/archive/master.zip)





### Source Code

If you’re a seasoned VB programmer, you’ve probably seen your share of VB programs that do modest things but do them in disturbingly complex ways. One of the most common abuses of this is the creation of classes to represent every last scrap of data. I’ve worked on projects that have literally dozens of classes defined just to represent the contents of modest databases. I find this clutter is usually pointless; where those dozens of classes can literally be replaced with one or two. And one nasty side effect of having these heaps of rubbish is that a simple change to the program can require a retooling of most of those classes, which certainly misses one of the central points of modularization.
<P>One simple yet effective way to abolish unnecessary classes is to use ad hoc data structures.
<P>An ad hoc data structure is a data structure that is created at run-time using some more general purpose data structure.
<P>What are some general purpose data structures we can use and how do we use them? One simple one is the array. We can use an array of variants to hold a simple data structure. Consider the following example of a data structure designed to represent a rectangle:
<UL><PRE>
<FONT COLOR="#000099">Begin Enum</FONT> RectProperties
    rLeft = 0
    rTop = 1
    rWidth = 2
    rHeight = 3
<FONT COLOR="#000099">End Enum</FONT>
<P><FONT COLOR="#000099">Private Sub</FONT> TrivialDemo
    <FONT COLOR="#000099">Dim</FONT> Rect(4) <FONT COLOR="#000099">As Variant</FONT>
    <FONT COLOR="#009900">'Populate the rectangle’s properties</FONT>
    Rect(rLeft) = 10
    Rect(rTop) = 10
    Rect(rWidth) = 100
    Rect(rHeight) = 50
    <FONT COLOR="#009900">'Use it for something</FONT>
    MsgBox Rect(rLeft)
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<P>Notice it’s a lot easier to manage a list of enumerated properties here than to create a whole class with properties, ad nauseum, just to represent a rectangle. It’s tempting to think a simple <TT>Type</TT> statement would be even easier to implement, but take note here that user-defined types cannot be public, which means they can’t readily be shared across classes, forms, ActiveX controls, etc. An array – or at least a Variant containing an array – can.
<P>What if we don’t want to deal with arrays and enumerations? One very good choice is the Collection. So long as you know the names of the properties you want to represent through some means external to the Collection object, you’ll have no problem dealing with it. Consider the same example code above, modified to use Collections.
<UL><PRE>
<P><FONT COLOR="#000099">Private Sub</FONT> TrivialDemo
    <FONT COLOR="#000099">Dim</FONT> Rect <FONT COLOR="#000099">As Collection</FONT>
    <FONT COLOR="#000099">Set</FONT> Rect = <FONT COLOR="#000099">New Collection</FONT>, RectCopy <FONT COLOR="#000099">As Collection</FONT>
    <FONT COLOR="#009900">'Populate the rectangle’s properties</FONT>
    Rect("Left") = 10
    Rect("Top") = 10
    Rect("Width") = 100
    Rect("Height") = 50
    <FONT COLOR="#009900">'Use it for something</FONT>
    MsgBox Rect("Left")
<FONT COLOR="#000099">End Sub</FONT>
</PRE></UL>
<P>If you’re willing to bring the Scripting runtime library into this, you can even use the venerable Dictionary object in a similar fashion. And there are still other options available, but these two will generally suffice for most simple data structures.
<P>But we don’t have to stop here. Generally, few data structures that matter go only one level deep like we’ve just shown. More commonly, a data structure has a number of single-value properties like in our examples and also a number of sets of other data structures. For example, a data structure representing an adult human might need a list of that person’s children. How do we do this sort of thing? The same way we’ve done up ‘til now, only we store Variant arrays or Collections (whichever the case may be) of another data structure – perhaps the same kind that holds the adult’s information. Using Collections, for example, you can find that accessing items in complicated data structures can be as straightforward as the following. Compare the Class way with the Collection way:
<UL><PRE>
MsgBox "One of my grandchildren’s names is " _
  Myself.Children(1).Children(1).Name
MsgBox "One of my grandchildren’s names is " _
  Myself("Children")(1)("Children")(1)("Name")
</PRE></UL>
<P>Note that for a few more characters and no less readability, we get to avoid the work involved in creating and maintaining a class.
<P>At this point, it might not seem like we’ve gained much, especially since we don’t have any simple way to tie methods or event triggers to our ad hoc data structures like we could with a class. But there are two enormous benefits ad hoc data structures can bestow: self definition and data definition.
Self definition refers to the idea of a data structure being able to represent a genera of other data structures. For example, if you program with ADO (or DAO or RDO), you know by now that VB doesn’t create a separate class for every table and another for every field. You get a small set of general-purpose data structures – connections, recordsets, field collections, and so on, and these all mold themselves to fit the particulars of whatever database elements they are connected with. Perhaps you hadn’t thought of them as such, but these classes are actually specialized ad hoc data structures.
<P>Following that model, you can create your own self-defining data structures. The first key is to define what is generally common among the genera of objects you want to model, to put all of those things in your class, and to leave out the properties (such as the names of fields in a table) particular to each instance. The second key is to find a way for this structure to define itself – those particular properties – based on information inherent in what’s being loaded. The ADO Recordset class, for instance, can find out about the fields in the tables it’s retrieving from the response from the database engine to its query. Many modern information servers can tell an entity querying it a lot about what it has to offer. This is what you target in your design. One of the greatest advantages of this approach is that there is generally little work involved in upgrading your self-defining data structures just because the properties particular instances change. Instead, you’ll be focussed primarily on dealing with the business end of your code, which will be the real target for changes, any way.
<P>A data-defined data structure is similar to a self-defined structure, except in principle, you are in charge of maintaining the definition. The definition could be stored in a text file, a database table, or even hard-coded in a definition module. The important key is that data apart from the actual code plays a central role in identifying what the various properties and collections of properties and so on will look like for a given instance of your ad hoc data structure. One of the incredible benefits of this approach is that it can be much easier to document the particular details of your application. There’s no reason, for instance, you couldn’t have this sort of information stored in Excel spreadsheets that your program can read and you can print out for reference documentation, to give a dramatic example. Change the definition to reflect a changing business need and you’ve got an instant upgrade of your documentation, too. And isn’t this sort of division of business rules from a flexible foundation an essential part of what you’re shooting for in the first place?
<P>In summary, ad hoc data structures offer strong flexibility and can simplify many of your projects – especially when used in conjunction with good class design. Further, ad hoc data structures can help you separate your business logic from your foundation code. And they can even be designed to morph into roles defined by the data they interface, saving you coding work, linking you to shifting standards, and shortening upgrade cycles. All this for a few more keystrokes.


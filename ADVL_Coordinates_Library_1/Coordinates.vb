'Public Class Coordinates
'The Coordinates classes are used to store geodetic and projected coordinate parameters and convert between coordinate systems.

'These classes are based on the tables in the EPSG database.
'Many of the classes correspond to tables in the database.
'For simplicity the data contained in multiple tables is often combined into a single class.

'References:
'Project \ Add Reference... \ ADVL_Utilties_Library_1

Public Enum FindIndexResults
        OK        'The specified record was found.
        NotFound  'The specified record was not found.
        ManyFound 'Many record match the record specifications. The first index value was returned.
        ListEmpty 'The list was empty. No recoreds to search.
    End Enum

    'Public Enum EnumCrsType
    Public Enum CrsTypes
        Compound
        Engineering
        Geocentric
        Geographic2D
        Geographic3D
        Projected
        Vertical
        Unknown
    End Enum



    Public Class CoordinateAxisName
        'Coordinate Axis Name class

        'NOTE: THE COORDINATE AXIS NAME TABLE ONLY EXISTS IN THE EPSG DATABASE FOR STORAGE EFFICIENCY, WITH REDUCED DUPLICATION.
        'TO SIMPLIFY THE DATA STRUCTURES IN THIS SOFTWARE, THIS DATA IS DUPLICATED IN THE CoordinateAxis CLASS.
        'The CoordinateAxisName class is only included here to display the contents of the Coordinate Axis Name table.

        'The Coordinate Axis Class stores the following information:
        '   Name
        '   Author
        '   Code
        '   Description
        '   Comments
        '   Deprecated

        'The name of this Coordinate Axis
        Private _name As String
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this Coordinate Axis Name.
        Private _author As String
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The code used by the author.
        Private _Code As Integer = 0
        Property Code As Integer
            Get
                Return _Code
            End Get
            Set(value As Integer)
                _Code = value
            End Set
        End Property

        'Description of this Coordinate Axis Name.
        Private _description As String
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        'Comments about this Coordinate Axis Name.
        Private _comments As String
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

    End Class


    Public Class CoordinateAxis
        'The Cooordinate Axis class stores the parameters of a coordinate axis.
        'One or more coordinate axes are used in a coordinate system.

        'The EPSG database uses separate tables for the Coordinate Axis and Coordinate Axis Name.
        '(Separate tables are used for efficiency of storage, with reduced duplication, but increases the data storage complexity.)
        'This class includes the parameters contained in the Coordinate Axis Name table.


        'The Axis class stores the following information:
        '   Author      The Author of the axis
        '   Code        The axis code defined by the author (The combined Author and Code fields are unique.)
        '   Name        The name of the axis (The name field may not be unique.)
        '   Description
        '   Comments
        '   Deprecated
        '   Orientation
        '   Abbreviation
        '   UnitOfMeasure   use clsUnitOfMeasure
        '   Order

        'A CoordinateSystemCode field is used in the EPSG database table to link the CoordinateAxis to the CoordinateSystem.
        'This is not required in the data stored in the Coordinates classes.
        'The CoordinateSystem class will contain the list of corresponding Coordinate Axes.

        'Alias names not required.

        Public UnitOfMeasure As New UnitOfMeasure

        'The author of this Coordinate Axis.
        Private _author As String
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The Coordinate Axis code defined by the author.
        Private _Code As Integer = 0
        Property Code As Integer
            Get
                Return _Code
            End Get
            Set(value As Integer)
                _Code = value
            End Set
        End Property

        'The name of the Axis (May not be unique - the same name may be used in an axis with a different orientation)
        Private _name As String
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'Description of this Coordinate Axis.
        Private _description As String
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        'Comments about the Coordinate Axis.
        Private _comments As String
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        'The orientation of the Coordinate Axis.
        'The direction of the positive increments of the axis: north, east, south, west, up, down,
        'or the approximate orientation in case of an oblique orientation: NE, SE, SW, NW.
        Private _orientation As String
        Property Orientation As String
            Get
                Return _orientation
            End Get
            Set(value As String)
                _orientation = value
            End Set
        End Property

        'The abbreviation of the Coordinate Axis.
        Private _abbreviation As String
        Property Abbreviation As String
            Get
                Return _abbreviation
            End Get
            Set(value As String)
                _abbreviation = value
            End Set
        End Property

        'The position of this axis within the Coordinate System: 1, 2 or 3.
        Private _order As Integer = 0
        Property Order As Integer
            Get
                Return _order
            End Get
            Set(value As Integer)
                _order = value
            End Set
        End Property

    End Class


    Public Class CoordinateSystem
        'Coordinate Systems class
        'This class stores the following information:
        '   Name        The name of the coordinate system
        '   Author      Author of the coordinate sstem
        '   Code        Code used by the Author for the coordinate system
        '   Selected
        '   Type
        '   Dimension   1, 2 or 3
        '   Comments
        '   Deprecated

        'Type of the Coordinate System: "affine", "Cartesian", "cylindrical", "ellipsoidal", "linear", "polar", "spherical" or "vertical". ("oblique Cartesian" and "gravity-related" have been replaced by "affine" and "vertical" respectively).
        'Public Enum EnumCSType
        Public Enum CSTypes
            Affine
            Cartesian
            Cylindrical
            Ellipsoidal
            Linear
            Polar
            Spherical
            Vertical
            Unknown
        End Enum

        Public AliasName As New List(Of String) 'Used to store alias names for this Coordinate System.
        Public Axis As New List(Of CoordinateAxis) 'Used to store the axes that form this coordinate system.

        'The name of the Coordinate System.
        Private _name As String
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this Coordinate System.
        Private _author As String
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The Coordinate System code defined by the author.
        Private _Code As Integer = 0
        Property Code As Integer
            Get
                Return _Code
            End Get
            Set(value As Integer)
                _Code = value
            End Set
        End Property

        'If True, this Coordinate System has been selected. This is used to process or save a subset of CSs in a list.
        Private _selected As Boolean = True
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'The type of Coordinate System (Affine, Cartesian, Cylindrical, Ellipsoidal, Linear, Polar, Spherical, Vertical, Unknown)
        Private _type As CSTypes
        Property Type As CSTypes
            Get
                Return _type
            End Get
            Set(value As CSTypes)
                _type = value
            End Set
        End Property

        'The number of dimensions of the Coordinate System: 1, 2 or 3.
        Private _dimension As Integer
        Property Dimension As Integer
            Get
                Return _dimension
            End Get
            Set(value As Integer)
                _dimension = value
            End Set
        End Property

        'Comments about the Coordinate System.
        Private _comments As String
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property


#Region " Methods" '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------


        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End Class


#Region "Coordinate Reference Systems" '-----------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class CoordinateReferenceSystemSummary
        'Coordinate Reference System summary data.
        'This class stores the following summary information for each Coordinate Reference System:
        '   Name        Name of the CRS
        '   Author      Author of the CRS
        '   Code        Code used by the Author for the CRS
        '   Area        Area of Use
        '   Type        Type of CRS
        '   Scope       Scope of the CRS
        '   Comments    
        '   Deprecated  "Yes" if CRS is deprecated, "No" if CRS is current and valid.
        '   
        'There are eight types of coordinate reference system:
        '   Compound
        '   Engineering
        '   Geocentric
        '   Geographic 2D
        '   Geographic 3D
        '   Projected
        '   Vertical

        'Separate classes are used to store the complete set of parameters of each type.


        ''Public Enum EnumCrsType
        'Public Enum CrsTypes
        '    Compound
        '    Engineering
        '    Geocentric
        '    Geographic2D
        '    Geographic3D
        '    Projected
        '    Vertical
        '    Unknown
        'End Enum

        Public AliasName As New List(Of String) 'Used to store alias names for this Coordinate Reference System.

        'Public Area As New clsAreaOfUse 'Used to store the area of use for this Coordinate Reference System.
        'Public CoordSystem As New clsCoordinateSystem 'Used to store the Coordinate System for this Coordinate Reference System.

        'Only Name, Author, Code and Type parameters are stored here for Area and Coordinate System.
        'The Author and Code values are sufficent to find the full set of parameters in the Area and Coordinate System lists.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New NameAuthorCodeInfo 'The coordinate system used by this Coordinate Reference System.

        'The name of the CRS
        Private _name As String
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this CRS summary.
        Private _author As String
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The CRS code defined by the author.
        Private _Code As Integer = 0
        Property Code As Integer
            Get
                Return _Code
            End Get
            Set(value As Integer)
                _Code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = True
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'The type of CRS (Compound, Engineering, Geocentric, Geographic 2D, Geographic 3D, Projected, Vertical)
        Private _type As CrsTypes
        Property Type As CrsTypes
            Get
                Return _type
            End Get
            Set(value As CrsTypes)
                _type = value
            End Set
        End Property

        'Scope of the CRS
        Private _scope
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments about the CRS.
        Private _comments As String
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified datum alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub

    End Class 'CoordinateReferenceSystemSummary

    Public Class HorizontalCRSSummary
        'The summary parameters of the horizontal coordinate reference system used in a Compound CRS.

        Public AliasName As New List(Of String) 'Used to store alias names for this Coordinate Reference System.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New NameAuthorCodeInfo 'The coordinate system used by this Coordinate Reference System.

        'The name of the CRS
        Private _name As String
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this CRS summary.
        Private _author As String
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The CRS code defined by the author.
        Private _Code As Integer = 0
        Property Code As Integer
            Get
                Return _Code
            End Get
            Set(value As Integer)
                _Code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = True
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'The type of CRS (Compound, Engineering, Geocentric, Geographic 2D, Geographic 3D, Projected, Vertical)
        Private _type As CrsTypes
        Property Type As CrsTypes
            Get
                Return _type
            End Get
            Set(value As CrsTypes)
                _type = value
            End Set
        End Property

        'Scope of the CRS
        Private _scope
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments about the CRS.
        Private _comments As String
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified datum alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub

    End Class

    Public Class CompoundCRSSummary
        'Compound coordinate reference system parameters.
        '

        Public AliasName As New List(Of String) 'Used to store alias names for this Coordinate Reference System.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use for this Coordinate Reference System.
        Public HorizontalCRS As New HorizontalCRSSummary
        Public VerticalCRS As New VerticalCRSSummary

        'The name of the CRS
        Private _name As String
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this CRS summary.
        Private _author As String
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The CRS code defined by the author.
        Private _Code As Integer = 0
        Property Code As Integer
            Get
                Return _Code
            End Get
            Set(value As Integer)
                _Code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = True
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'NOTE: The CRS type is Compound!!!
        ''The type of CRS (Compound, Engineering, Geocentric, Geographic 2D, Geographic 3D, Projected, Vertical)
        'Private _type As CrsTypes
        'Property Type As CrsTypes
        '    Get
        '        Return _type
        '    End Get
        '    Set(value As CrsTypes)
        '        _type = value
        '    End Set
        'End Property

        'Scope of the CRS
        Private _scope
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments about the CRS.
        Private _comments As String
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified datum alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub


    End Class

    Public Class EngineeringCRS
        'Compound coordinate reference system parameters.
        '(This includes the Coordinate Reference System Summary data.)

    End Class

    Public Class GeocentricCRS
        'Compound coordinate reference system parameters.
        '(This includes the Coordinate Reference System Summary data.)

    End Class

    Public Class GeographicCRSSummary
        'A summary of the parameters in a Geographic Coordinate Reference System.
        'This is used to decribe the Source Geographic Coordinate Reference System use in some of the 2D Geographic Coordinate Reference Systems.

        'Public Enum EnumCrsType
        Public Enum CrsTypes
            Compound
            Engineering
            Geocentric
            Geographic2D
            Geographic3D
            Projected
            Vertical
            Unknown
        End Enum

        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Reference System.

        'The name of the Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'The type of CRS (Compound, Engineering, Geocentric, Geographic 2D, Geographic 3D, Projected, Vertical)
        Private _type As CrsTypes
        Property Type As CrsTypes
            Get
                Return _type
            End Get
            Set(value As CrsTypes)
                _type = value
            End Set
        End Property

        'The scope of the Coordinate Reference System.
        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments on the Coordinate Reference System.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate Reference System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate Reference System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'GeographicCRSSummary

    Public Class Geographic2DCRSSummary
        'Summary Parameters for a Geographic 2D Coordinate Reference System.
        'Excludes the detailed parameters for the Area of Use, Coordinate System, Datum and Base Coordinate Reference System.

        'Summary Parameters:
        '   Name        Name of the CRS
        '   Author      Author of the CRS
        '   Code        Code used by the Author for the CRS
        '   Scope       Scope of the CRS
        '   Comments    Comments on the CRS.
        '   Deprecated  "Yes" if CRS is deprecated, "No" if CRS is current and valid.

        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Reference System.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New NameAuthorCodeInfo 'The coordinate system used by this Coordinate Reference System.
        Public Datum As New NameAuthorCodeInfo 'A summary of the Datum used by this Coordinate Reference System.
        Public SourceGeographicCRS As New NameAuthorCodeInfo 'A summary of the Source Geographic CRS used by this Coordinate Reference System


        'The name of the Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments on the Coordinate Reference System.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate Reference System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate Reference System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'Geographic2DCRSSummary


    Public Class Geographic3DCRSSummary
        'Summary Parameters for a Geographic 3D Coordinate Reference System.
        'Excludes the detailed parameters for the Area of Use, Coordinate System, Datum and Base Coordinate Reference System.

        'Summary Parameters:
        '   Name        Name of the CRS
        '   Author      Author of the CRS
        '   Code        Code used by the Author for the CRS
        '   Scope       Scope of the CRS
        '   Comments    Comments on the CRS.
        '   Deprecated  "Yes" if CRS is deprecated, "No" if CRS is current and valid.

        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Reference System.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New NameAuthorCodeInfo 'The coordinate system used by this Coordinate Reference System.
        Public Datum As New NameAuthorCodeInfo 'A summary of the Datum used by this Coordinate Reference System.
        Public SourceGeographicCRS As New NameAuthorCodeInfo 'A summary of the Source Geographic CRS used by this Coordinate Reference System


        'The name of the Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments on the Coordinate Reference System.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate Reference System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate Reference System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'Geographic3DCRSSummary

    Public Class GeocentricCRSSummary
        'Summary parameters for a Geocentric Coordinate Reference System.
        'Excludes the detailed parameters for the Area of Use, Coordinate System and Datum.

        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Reference System.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New NameAuthorCodeInfo 'The coordinate system used by this Coordinate Reference System.
        Public Datum As New NameAuthorCodeInfo 'A summary of the Datum used by this Coordinate Reference System.
        'Public SourceGeographicCRS As New NameAuthorCodeInfo 'A summary of the Source Geographic CRS used by this Coordinate Reference System

        'The name of the Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments on the Coordinate Reference System.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate Reference System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate Reference System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'GeocentricCRSSummary

    Public Class VerticalCRSSummary
        'Summary parameters for a Vertical Coordinate Reference System.
        'Excludes the detailed parameters for the Area of Use, Coordinate System and Datum.

        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Reference System.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New NameAuthorCodeInfo 'The coordinate system used by this Coordinate Reference System.
        Public Datum As New NameAuthorCodeInfo 'A summary of the Datum used by this Coordinate Reference System.
        'Public SourceGeographicCRS As New NameAuthorCodeInfo 'A summary of the Source Geographic CRS used by this Coordinate Reference System

        'The name of the Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments on the Coordinate Reference System.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate Reference System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate Reference System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'VerticalCRSSummary

    Public Class EngineeringCRSSummary
        'Summary parameters for an Engineering Coordinate Reference System.
        'Excludes the detailed parameters for the Area of Use, Coordinate System and Datum.

        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Reference System.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New NameAuthorCodeInfo 'The coordinate system used by this Coordinate Reference System.
        Public Datum As New NameAuthorCodeInfo 'A summary of the Datum used by this Coordinate Reference System.

        'The name of the Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments on the Coordinate Reference System.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate Reference System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate Reference System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'EngineeringCRSSummary

    Public Class NameAuthorCodeInfo
        'Summary information used to identify a particular parameter set (such as an Area of Use or a Datum).

        'The name of the Parameter Set.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Parameter Set.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Parameter Set.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'The type of Parameter Set.
        Private _type As String = ""
        Property Type As String
            Get
                Return _type
            End Get
            Set(value As String)
                _type = value
            End Set
        End Property

    End Class 'NameAuthorCodeInfo

    Public Class Geographic2DCRS
        'Parameters for a Geographic 2D Coordinate Reference System.
        '(This includes the Coordinate Reference System Summary data.)
        'NOTE: The Geographic2DCRSSummary is used in the Geographic2DCRSList to avoid duplicating Area, CoordinateSystem, Datum and SourceGeographicCRS data.
        '      The Geographic2DCRSSummary includes only the Name, Author and Code used to identify the complete record of parameters in the corresponding list.

        'Summary Parameters:
        '   Name        Name of the CRS
        '   Author      Author of the CRS
        '   Code        Code used by the Author for the CRS
        '   Area        Area of Use
        '   Type        Type of CRS
        '   Scope       Scope of the CRS
        '   Comments    Comments on the CRS.
        '   Deprecated  "Yes" if CRS is deprecated, "No" if CRS is current and valid.

        'Geographic 2D CRS Parameters:
        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Reference System.
        Public Area As New AreaOfUse 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New CoordinateSystem 'The coordinate system used by this Coordinate Reference System.
        Public Datum As New DatumSummary 'A summary of the Datum used by this Coordinate Reference System.
        Public SourceGeographicCRS As New GeographicCRSSummary 'A summary of the Source Geographic CRS used by this Coordinate Reference System


        'The name of the Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments on the Coordinate Reference System.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate Reference System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate Reference System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class

    Public Class Geographic3DCRS
        'Geographic coordinate reference system parameters.
        '(This includes the Coordinate Reference System Summary data.)
        'NOTE: The Geographic3DCRSSummary is used in the Geographic2DCRSList to avoid duplicating Area, CoordinateSystem, Datum and SourceGeographicCRS data.
        '      The Geographic3DCRSSummary includes only the Name, Author and Code used to identify the complete record of parameters in the corresponding list.

        'Summary Parameters:
        '   Name        Name of the CRS
        '   Author      Author of the CRS
        '   Code        Code used by the Author for the CRS
        '   Area        Area of Use
        '   Type        Type of CRS
        '   Scope       Scope of the CRS
        '   Comments    Comments on the CRS.
        '   Deprecated  "Yes" if CRS is deprecated, "No" if CRS is current and valid.

        'Geographic 3D CRS Parameters:
        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Reference System.
        Public Area As New AreaOfUse 'Used to store the area of use for this Coordinate Reference System.
        Public CoordinateSystem As New CoordinateSystem 'The coordinate system used by this Coordinate Reference System.
        Public Datum As New DatumSummary 'A summary of the Datum used by this Coordinate Reference System.
        Public SourceGeographicCRS As New GeographicCRSSummary 'A summary of the Source Geographic CRS used by this Coordinate Reference System


        'The name of the Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this CRS has been selected. This is used to process or save a subset of CRSs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments on the Coordinate Reference System.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Coordinate Reference System alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Coordinate Reference System alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class

    Public Class ProjectedCRS
        'Projected coordinate reference system parameters.
        '(This includes the Coordinate Reference System Summary data.)

        'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
        'These fields in the Coordinate Reference System table are not used for projected CRSs: DATUM_CODE CMPD_HORIZCRS_CODE CMPD_VERTCRS_CODE
        'For Projected CRS, COORD_REF_SYS_KIND = projected.

        'PROPERTIES:
        'Name
        'Author
        'Code
        'Selected
        'Scope
        'Comments
        'Deprecated

        Public AliasName As New List(Of String) 'Used to store alias names for the Projection.
        Public Area As New NameAuthorCodeInfo 'Used to store the Area of Use identifier for this Projected CRS. 'This is used to identify the full set of parameters in the Area Of Use list.
        Public CoordinateSystem As New NameAuthorCodeInfo 'Used to store the Coordinate System identifier for this Projected CRS.
        Public SourceGeographicCRS As New NameAuthorCodeInfo 'Used to store the Source Geographic Coordinate Reference System identifier for this Projected CRS.
        Public Projection As New NameAuthorCodeInfo 'Used to store the Projection Operation identifier for this Projected CRS. (Coordinate_Operation, COORD_OP_TYPE = conversion) (Examples: Australian Map Grid zone 51, Map Grid of Australia zone 50)
        Public ProjectionMethod As New NameAuthorCodeInfo 'Used to store the Projection Method identifier for this Projected CRS. (Coordinate_Operation Method) (Examples: Transverse Mercator, Cassini-Soldner)

        'The name of the Projected CRS.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Projected CRS.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Projected CRS.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Projected CRS has been selected. This is used to process or save a subset of Projected CRS in a list.
        Private _selected As Boolean = True
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Comments on the Projected CRS.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'The scope of the Projected CRS
        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Projected CRS alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Projected CRS alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class


    'NOTE: All projected Coordinate Reference Systems are now stored in the one class: ProjectedCRS


    Public Class VerticalCRS
        'Vertical coordinate reference system parameters.
        '(This includes the Coordinate Reference System Summary data.)

    End Class

#End Region 'Coordinate Reference Systems ---------------------------------------------------------------------------------------------------------------------------------------------------




    Public Class TransverseMercator
        'The Transverse Mercator class stores the projection parameters and converts coordinates between the geographic and projected values.

        'PROPERTIES:
        'Projection
        '   Name
        '   Author
        '   Code
        '   LatitudeOfNaturalOrigin
        '   LongitudeOfNaturalOrigin
        '   ScaleFactorAtNaturalOrigin
        '   FalseEasting
        '   FalseNorthing

        'GeographicCRS
        '   Name
        '   Author
        '   Code
        '   EllipsoidName
        '   MajorSemiAxis
        '   InverseFlattening

        'Location
        '   Latitude        The geodetic Latitude coordinate
        '   Longitude       The geodetic Longitude coordinate
        '   EllipsoidHeight
        '   X               The cartesian X coordinate
        '   Y               The cartesian Y coordinate
        '   Z               The cartesian Z coordinate
        '   Easting         The projected Easting coordinate
        '   Northing        The projected Northing coordinate

        'METHODS:
        'ConvertLatLongToEastNorth
        'ConvertEastNorthToLatLong
        'Clear - Set all the parameter values to Double.NaN
        'CheckParameters - Check if valid parameter values have been set.

        Public Projection As New TransverseMercatorProjectionParameters
        Public GeographicCRS As New GeographicCRSParameters
        Public Location As New Location

        'Checks that all required parameter values have been entered.
        Private _parametersReady As Boolean = False
        ReadOnly Property ParametersReady As Boolean
            Get
                CheckParameters()
                Return _parametersReady
            End Get
        End Property

        'String describing the status of the projection parameters.
        'This is only updated when the ParametersReady property is read.
        Private _parameterStatus As String
        ReadOnly Property ParameterStatus As String
            Get
                Return _parameterStatus
            End Get
        End Property

#Region " Transverse Mercator Methods" '-----------------------------------------------------------------------------------------------------------------------------------------------------

        Public Sub Clear()
            'Clear all properties

            Projection.Name = ""
            Projection.Author = ""
            Projection.Code = 0
            Projection.LatitudeOfNaturalOrigin = Double.NaN
            Projection.LongitudeOfNaturalOrigin = Double.NaN
            Projection.ScaleFactorAtNaturalOrigin = Double.NaN
            Projection.FalseEasting = Double.NaN
            Projection.FalseNorthing = Double.NaN

            GeographicCRS.Name = ""
            GeographicCRS.Author = ""
            GeographicCRS.DatumName = ""
            GeographicCRS.EllipsoidName = ""
            GeographicCRS.SemiMajorAxis = Double.NaN
            GeographicCRS.InverseFlattening = Double.NaN
            GeographicCRS.SemiMinorAxis = Double.NaN

            Location.Latitude = Double.NaN
            Location.Longitude = Double.NaN
            Location.EllipsoidalHeight = Double.NaN
            Location.X = Double.NaN
            Location.Y = Double.NaN
            Location.Z = Double.NaN
            Location.Easting = Double.NaN
            Location.Northing = Double.NaN
        End Sub

        Private Sub CheckParameters()
            'Check if the parameters have been entered ready for projection calculations.

            Dim ErrorMessage As String = ""
            Dim ParameterError As Boolean = False

            If Projection.LatitudeOfNaturalOrigin = Double.NaN Then
                ParameterError = True
                ErrorMessage = "No latitude of natural origin entered."
            End If

            If Projection.LongitudeOfNaturalOrigin = Double.NaN Then
                ParameterError = True
                If ErrorMessage = "" Then
                    ErrorMessage = "No longitude of natural origin entered."
                Else
                    ErrorMessage = ErrorMessage & vbCrLf & "No longitude of natural origin entered."
                End If
            End If

            If Projection.ScaleFactorAtNaturalOrigin = Double.NaN Then
                ParameterError = True
                If ErrorMessage = "" Then
                    ErrorMessage = "No scale factor at natural origin entered."
                Else
                    ErrorMessage = ErrorMessage & vbCrLf & "No scale factor at natural origin entered."
                End If
            End If

            If Projection.FalseEasting = Double.NaN Then
                ParameterError = True
                If ErrorMessage = "" Then
                    ErrorMessage = "No false easting entered."
                Else
                    ErrorMessage = ErrorMessage & vbCrLf & "No false easting entered."
                End If
            End If

            If Projection.FalseNorthing = Double.NaN Then
                ParameterError = True
                If ErrorMessage = "" Then
                    ErrorMessage = "No false northing entered."
                Else
                    ErrorMessage = ErrorMessage & vbCrLf & "No false northing entered."
                End If
            End If

            If GeographicCRS.SemiMajorAxis = Double.NaN Then
                ParameterError = True
                If ErrorMessage = "" Then
                    ErrorMessage = "No semi major axis entered."
                Else
                    ErrorMessage = ErrorMessage & vbCrLf & "No semi major axis entered."
                End If
            End If

            If GeographicCRS.InverseFlattening = Double.NaN Then
                ParameterError = True
                If ErrorMessage = "" Then
                    ErrorMessage = "No semi minor axis entered."
                Else
                    ErrorMessage = ErrorMessage & vbCrLf & "No semi minor axis entered."
                End If
            End If

            If ParameterError = True Then
                _parametersReady = False
                _parameterStatus = ErrorMessage
            Else
                _parametersReady = True
                _parameterStatus = "Parameters OK"
            End If

        End Sub

        Public Sub LatLonToXYZ()
            'Calculates the cartesian X,Y,Z coordinates from the Latitude and Longitude values.

            Dim Flattening As Double
            Dim EccSquared As Double
            Dim CurvatureV As Double
            Dim LatitudeRadians As Double
            Dim LongitudeRadians As Double
            Dim Pi As Double

            Pi = 3.14159265358979
            Flattening = 1 / GeographicCRS.InverseFlattening
            EccSquared = 2 * Flattening - Flattening * Flattening
            LatitudeRadians = 2 * Pi * Location.Latitude / 360
            LongitudeRadians = 2 * Pi * Location.Longitude / 360
            CurvatureV = GeographicCRS.SemiMajorAxis / System.Math.Sqrt(1 - EccSquared * System.Math.Sin(LatitudeRadians) * System.Math.Sin(LatitudeRadians))

            Location.X = (CurvatureV + Location.EllipsoidalHeight) * System.Math.Cos(LatitudeRadians) * System.Math.Cos(LongitudeRadians)
            Location.Y = (CurvatureV + Location.EllipsoidalHeight) * System.Math.Cos(LatitudeRadians) * System.Math.Sin(LongitudeRadians)
            Location.Z = ((1 - EccSquared) * CurvatureV + Location.EllipsoidalHeight) * System.Math.Sin(LatitudeRadians)

        End Sub

        Public Sub XYZToLatLonHt()
            'Calculates the latitude, longitude and height values from the cartesian X,Y,Z coordinate values.

            Dim P As Double
            Dim R As Double
            Dim Mu As Double
            Dim Pi As Double
            Dim Flattening As Double
            Dim EccSquared As Double
            Dim CurvatureV As Double
            Dim LongitudeRadians As Double
            Dim LatitudeRadians As Double
            Dim LatTop As Double
            Dim LatBottom As Double

            Pi = 3.14159265358979
            Flattening = 1.0# / GeographicCRS.InverseFlattening
            EccSquared = 2 * Flattening - Flattening * Flattening

            P = System.Math.Sqrt(Location.X * Location.X + Location.Y * Location.Y)
            R = System.Math.Sqrt(P * P + Location.Z * Location.Z)

            If R = 0 Then
                Mu = 0
                'If blnShowMessages = True Then 'Show the error message
                '    RaiseEvent Warning("Parameter R = 0. Result may be in error.")
                'End If
            Else
                Mu = System.Math.Atan((Location.Z / P) * ((1 - Flattening) + (EccSquared * GeographicCRS.SemiMajorAxis) / R))
            End If

            If Location.X = 0 Then
                LongitudeRadians = 1.5707963267949
                'If blnShowMessages = True Then 'Show the error message
                '    RaiseEvent Warning("Datum1Locn.X = 0. Result may be in error.")
                'End If
            Else
                LongitudeRadians = System.Math.Atan(Location.Y / Location.X)
            End If

            LatTop = Location.Z * (1 - Flattening) + EccSquared * GeographicCRS.SemiMajorAxis * System.Math.Sin(Mu) * System.Math.Sin(Mu) * System.Math.Sin(Mu)
            LatBottom = (1 - Flattening) * (P - EccSquared * GeographicCRS.SemiMajorAxis * System.Math.Cos(Mu) * System.Math.Cos(Mu) * System.Math.Cos(Mu))
            LatitudeRadians = System.Math.Atan(LatTop / LatBottom)
            Location.EllipsoidalHeight = P * System.Math.Cos(LatitudeRadians) + Location.Z * System.Math.Sin(LatitudeRadians) - GeographicCRS.SemiMajorAxis * System.Math.Sqrt(1 - EccSquared * System.Math.Sin(LatitudeRadians) * System.Math.Sin(LatitudeRadians))

            If LongitudeRadians < 0 Then LongitudeRadians = LongitudeRadians + Pi
            Location.Longitude = (LongitudeRadians / Pi) * 180
            Location.Latitude = (LatitudeRadians / Pi) * 180

        End Sub

        Public Sub LatLonToEastNorth()
            'Converts geographic Latitude and Longitude coordinates to projected Easting and Northing coordinates.

            Dim LatRad As Double 'Latitude in radians
            Dim Mdist As Double 'Meridian distance
            Dim A0 As Double
            Dim A2 As Double
            Dim A4 As Double
            Dim A6 As Double
            Dim Flattening As Double
            Dim E2 As Double 'Eccentricity squared
            Dim E4 As Double
            Dim E6 As Double
            Dim a As Double 'Semi major axis
            Dim Term1 As Double
            Dim Term2 As Double
            Dim Term3 As Double
            Dim Term4 As Double
            Dim Pi As Double
            Dim SinLat As Double
            Dim Sin2Lat As Double
            Dim Sin4Lat As Double
            Dim Sin6Lat As Double
            Dim Rho As Double
            Dim Nu As Double
            Dim CosLat As Double
            Dim CosLat2 As Double
            Dim CosLat3 As Double
            Dim CosLat4 As Double
            Dim CosLat5 As Double
            Dim CosLat6 As Double
            Dim CosLat7 As Double
            Dim CosLat8 As Double
            Dim DifLonRad As Double
            Dim DifLonRad2 As Double
            Dim DifLonRad3 As Double
            Dim DifLonRad4 As Double
            Dim DifLonRad5 As Double
            Dim DifLonRad6 As Double
            Dim DifLonRad7 As Double
            Dim DifLonRad8 As Double
            Dim Psi As Double
            Dim Psi2 As Double
            Dim Psi3 As Double
            Dim Psi4 As Double
            Dim TanLat As Double
            Dim TanLat2 As Double
            Dim TanLat4 As Double
            Dim TanLat6 As Double
            Dim GridConv As Double 'Grid Convergence
            Dim PointScale As Double 'Point Scale
            Dim InvalidParams As Boolean

            Pi = 3.14159265358979
            LatRad = (Location.Latitude / 180) * Pi
            Flattening = 1.0# / GeographicCRS.InverseFlattening
            E2 = (2.0# * Flattening) - (Flattening * Flattening)
            E4 = E2 * E2
            E6 = E2 * E4
            A0 = 1.0# - (E2 / 4.0#) - ((3.0# * E4) / 64.0#) - ((5.0# * E6) / 256.0#)
            A2 = (3.0# / 8.0#) * (E2 + (E4 / 4.0#) + ((15.0# * E6) / 128.0#))
            A4 = (15.0# / 256.0#) * (E4 + ((3.0# * E6) / 4.0#))
            A6 = (35.0# * E6) / 3072.0#
            SinLat = System.Math.Sin(LatRad)
            Sin2Lat = System.Math.Sin(2.0# * LatRad)
            Sin4Lat = System.Math.Sin(4.0# * LatRad)
            Sin6Lat = System.Math.Sin(6.0# * LatRad)
            a = GeographicCRS.SemiMajorAxis
            Term1 = a * A0 * LatRad
            Term2 = -a * A2 * Sin2Lat
            Term3 = a * A4 * Sin4Lat
            Term4 = -a * A6 * Sin6Lat
            Mdist = Term1 + Term2 + Term3 + Term4 'The meridian distance.

            Rho = a * (1.0# - E2) / (1.0# - (E2 * SinLat * SinLat)) ^ 1.5
            Nu = a / (1.0# - (E2 * SinLat * SinLat)) ^ 0.5

            DifLonRad = ((Location.Longitude - Projection.LongitudeOfNaturalOrigin) / 180) * Pi
            DifLonRad2 = DifLonRad * DifLonRad
            DifLonRad3 = DifLonRad2 * DifLonRad
            DifLonRad4 = DifLonRad2 * DifLonRad2
            DifLonRad5 = DifLonRad4 * DifLonRad
            DifLonRad6 = DifLonRad3 * DifLonRad3
            DifLonRad7 = DifLonRad3 * DifLonRad4
            DifLonRad8 = DifLonRad4 * DifLonRad4

            CosLat = System.Math.Cos(LatRad)
            CosLat2 = CosLat * CosLat
            CosLat3 = CosLat2 * CosLat
            CosLat4 = CosLat2 * CosLat2
            CosLat5 = CosLat3 * CosLat2
            CosLat6 = CosLat3 * CosLat3
            CosLat7 = CosLat4 * CosLat3
            CosLat8 = CosLat4 * CosLat4

            Psi = Nu / Rho
            Psi2 = Psi * Psi
            Psi3 = Psi2 * Psi
            Psi4 = Psi2 * Psi2

            TanLat = System.Math.Tan(LatRad)
            TanLat2 = TanLat * TanLat
            TanLat4 = TanLat2 * TanLat2
            TanLat6 = TanLat4 * TanLat2

            Term1 = Nu * DifLonRad * CosLat
            Term2 = Nu * DifLonRad3 * CosLat3 * (Psi - TanLat2) / 6.0#
            Term3 = Nu * DifLonRad5 * CosLat5 * (4.0# * Psi3 * (1.0# - 6.0# * TanLat2) + Psi2 * (1.0# + 8.0# * TanLat2) - Psi * (2.0# * TanLat2) + TanLat4) / 120.0#
            Term4 = Nu * DifLonRad7 * CosLat7 * (61.0# - 479.0# * TanLat2 * 179.0# * TanLat4 - TanLat6) / 5040.0#
            Location.Easting = Term1 + Term2 + Term3 + Term4
            Location.Easting = Location.Easting * Projection.ScaleFactorAtNaturalOrigin
            Location.Easting = Location.Easting + Projection.FalseEasting 'Easting value

            Term1 = Nu * SinLat * DifLonRad2 * CosLat / 2.0#
            Term2 = Nu * SinLat * DifLonRad4 * CosLat3 * (4.0# * Psi2 + Psi - TanLat2) / 24.0#
            Term3 = Nu * SinLat * DifLonRad6 * CosLat5 * (8.0# * Psi4 * (11.0# - 24.0# * TanLat2) - 28.0# * Psi3 * (1.0# - 6.0# * TanLat2) + Psi2 * (1.0# - 32.0# * TanLat2) - Psi * (2.0# * TanLat2) + TanLat4) / 720.0#
            Term4 = Nu * SinLat * DifLonRad8 * CosLat7 * (1385.0# - 3111.0# * TanLat2 + 543.0# * TanLat4 - TanLat6) / 40320.0#
            Location.Northing = Mdist + Term1 + Term2 + Term3 + Term4
            Location.Northing = Location.Northing * Projection.ScaleFactorAtNaturalOrigin
            Location.Northing = Location.Northing + Projection.FalseNorthing 'Northing value

            'Calculate grid convergence:
            Term1 = -SinLat * DifLonRad
            Term2 = -SinLat * DifLonRad3 * CosLat2 * (2.0# * Psi2 - Psi) / 3.0#
            Term3 = -SinLat * DifLonRad5 * CosLat4 * (Psi4 * (11.0# - 24.0# * TanLat2) - Psi3 * (11.0# - 36.0# * TanLat2) + 2.0# * Psi2 * (1.0# - 7.0# * TanLat2) + Psi * TanLat2) / 15.0#
            Term4 = SinLat * DifLonRad7 * CosLat6 * (17.0# - 26.0# * TanLat2 + 2.0# * TanLat4) / 315.0#
            GridConv = Term1 + Term2 + Term3 + Term4
            GridConv = (GridConv / Pi) * 180 'Convert to degrees

            'Calculate point scale:
            Term1 = 1.0# + (DifLonRad2 * CosLat2 * Psi) / 2.0#
            Term2 = DifLonRad4 * CosLat4 * (4.0# * Psi3 * (1.0# - 6.0# * TanLat2) + Psi2 * (1.0# + 24.0# * TanLat2) - 4.0# * Psi * TanLat2) / 24.0#
            Term3 = DifLonRad6 * CosLat6 * (61.0# - 148.0# * TanLat2 + 16.0# * TanLat4) / 720.0#
            PointScale = Term1 + Term2 + Term3

        End Sub

        Public Sub EastNorthToLatLon()
            'Converts projected Easting and Northing coordinates to geographic Latitude and Longitude coordinates.

            Dim Pi As Double
            Dim Edash As Double
            Dim EdashOnK0 As Double
            Dim Ndash As Double
            Dim m As Double
            Dim Sigma As Double
            Dim G As Double
            Dim N As Double
            Dim N2 As Double
            Dim N3 As Double
            Dim N4 As Double
            Dim a As Double 'Semi major axis
            Dim b As Double 'Semi-minor axis
            Dim Flattening As Double
            Dim Term1 As Double
            Dim Term2 As Double
            Dim Term3 As Double
            Dim Term4 As Double
            Dim Term5 As Double
            Dim FPLat As Double 'Foot Point Latitude
            Dim SinFPLat As Double
            Dim SecFPLat As Double
            Dim TanFPLat As Double
            Dim TanFPLat2 As Double
            Dim TanFPLat3 As Double
            Dim TanFPLat4 As Double
            Dim TanFPLat5 As Double
            Dim TanFPLat6 As Double
            Dim E2 As Double
            Dim Rho As Double
            Dim Nu As Double
            Dim Psi As Double 'Nu / Rho
            Dim Psi2 As Double
            Dim Psi3 As Double
            Dim Psi4 As Double
            Dim TonK0NuRho As Double
            Dim K0 As Double
            Dim X As Double
            Dim X2 As Double
            Dim X3 As Double
            Dim X4 As Double
            Dim X5 As Double
            Dim X6 As Double
            Dim X7 As Double
            Dim E2onK02NuRho As Double
            Dim E2onK02NuRho2 As Double
            Dim E2onK02NuRho3 As Double
            Dim CMRad As Double
            Dim GridConv As Double
            Dim PointScale As Double
            Dim InvalidParams As Boolean

            Pi = 3.14159265358979
            K0 = Projection.ScaleFactorAtNaturalOrigin

            Edash = Location.Easting - Projection.FalseEasting
            EdashOnK0 = Edash / K0

            Ndash = Location.Northing - Projection.FalseNorthing
            m = Ndash / K0
            Flattening = 1.0# / GeographicCRS.InverseFlattening
            a = GeographicCRS.SemiMajorAxis
            b = a * (1 - Flattening)
            N = (a - b) / (a + b)
            N2 = N * N
            N3 = N2 * N
            N4 = N2 * N2
            G = a * (1.0# - N) * (1 - N2) * (1.0# + (9.0# * N2) / 4.0# + (225.0# * N4) / 64.0#) * Pi / 180.0#
            Sigma = (m * Pi) / (G * 180.0#)

            Term1 = Sigma
            Term2 = ((3.0# * N / 2.0#) - (27.0# * N3 / 32.0#)) * System.Math.Sin(Sigma * 2.0#)
            Term3 = ((21.0# * N2 / 16.0#) - (55.0# * N4 / 32.0#)) * System.Math.Sin(Sigma * 4.0#)
            Term4 = (151.0# * N3) * System.Math.Sin(Sigma * 6.0#) / 96.0#
            Term5 = 1097.0# * N4 * System.Math.Sin(Sigma * 8.0#) / 512.0#
            FPLat = Term1 + Term2 + Term3 + Term4 + Term5 'The Foot point latitude.

            SinFPLat = System.Math.Sin(FPLat)
            SecFPLat = 1.0# / System.Math.Cos(FPLat)

            TanFPLat = System.Math.Tan(FPLat)
            TanFPLat2 = TanFPLat * TanFPLat
            TanFPLat3 = TanFPLat2 * TanFPLat
            TanFPLat4 = TanFPLat2 * TanFPLat2
            TanFPLat5 = TanFPLat3 * TanFPLat2
            TanFPLat6 = TanFPLat3 * TanFPLat3

            E2 = (2 * Flattening) - (Flattening * Flattening)
            Rho = a * (1.0# - E2) / (1.0# - E2 * SinFPLat * SinFPLat) ^ 1.5
            Nu = a / (1.0# - E2 * SinFPLat * SinFPLat) ^ 0.5

            TonK0NuRho = TanFPLat / (K0 * Nu)

            X = EdashOnK0 / Nu
            X2 = X * X
            X3 = X2 * X
            X4 = X2 * X2
            X5 = X3 * X2
            X6 = X3 * X3
            X7 = X4 * X3

            E2onK02NuRho = (EdashOnK0 * EdashOnK0) / (Rho * Nu)
            E2onK02NuRho2 = E2onK02NuRho * E2onK02NuRho
            E2onK02NuRho3 = E2onK02NuRho2 * E2onK02NuRho

            Psi = Nu / Rho
            Psi2 = Psi * Psi
            Psi3 = Psi2 * Psi
            Psi4 = Psi2 * Psi2

            Term1 = -((TanFPLat / (K0 * Rho)) * X * Edash / 2.0#)
            Term2 = (TanFPLat / (K0 * Rho)) * (X3 * Edash / 24.0#) * (-4.0# * Psi2 + 9.0# * Psi * (1.0# - TanFPLat2) + 12.0# * TanFPLat2)
            Term3 = -(TanFPLat / (K0 * Rho)) * (X5 * Edash / 720.0#) * (8.0# * Psi4 * (11.0# - 24.0# * TanFPLat2) - 12.0# * Psi3 * (21.0# - 71.0# * TanFPLat2) + 15.0# * Psi2 * (15.0# - 98.0# * TanFPLat2 + 15.0# * TanFPLat4) + 180.0# * Psi * (5.0# * TanFPLat2 - 3.0# * TanFPLat4) + 360.0# * TanFPLat4)
            Term4 = (TanFPLat / (K0 * Rho)) * (X7 * Edash / 40320.0#) * (1385.0# + 3633.0# * TanFPLat2 + 4095.0# * TanFPLat4 + 1575.0# * TanFPLat6)
            Location.Latitude = FPLat + Term1 + Term2 + Term3 + Term4
            Location.Latitude = (Location.Latitude / Pi) * 180 'Final latitude converted to degrees.

            CMRad = (Projection.LongitudeOfNaturalOrigin / 180) * Pi 'Central Meridian in radians
            Term1 = SecFPLat * X
            Term2 = -SecFPLat * (X3 / 6.0#) * (Psi + 2.0# * TanFPLat2)
            Term3 = SecFPLat * (X5 / 120.0#) * (-4.0# * Psi3 * (1.0# - 6.0# * TanFPLat2) + Psi2 * (9.0# - 68.0# * TanFPLat2) + 72.0# * Psi * TanFPLat2 + 24.0# * TanFPLat4)
            Term4 = -SecFPLat * (X7 / 5040.0#) * (61.0# + 662.0# * TanFPLat2 + 1320.0# * TanFPLat4 + 720.0# * TanFPLat6)
            Location.Longitude = CMRad + Term1 + Term2 + Term3 + Term4
            'Longitude = (Longitude / Pi) * 180 'Final longitude converted to degrees
            Location.Longitude = (Location.Longitude / Pi) * 180 'Final longitude converted to degrees

            'Calculate grid convergence:
            Term1 = -TanFPLat * X
            Term2 = (TanFPLat * X3 / 3.0#) * (-2.0# * Psi2 + 3.0# * Psi + TanFPLat2)
            Term3 = -(TanFPLat * X5 / 15.0#) * (Psi4 * (11.0# - 24.0# * TanFPLat2) - 3.0# * Psi3 * (8.0# - 23.0# * TanFPLat2) + 5.0# * Psi2 * (3.0# - 14.0# * TanFPLat2) + 30.0# * Psi * TanFPLat2 + 3.0# * TanFPLat4)
            Term4 = (TanFPLat * X7 / 315.0#) * (17.0# + 77.0# * TanFPLat2 + 105.0# * TanFPLat4 + 45.0# * TanFPLat6)
            GridConv = Term1 + Term2 + Term3 + Term4
            GridConv = (GridConv / Pi) * 180 'Convert to degrees

            'Calculate point scale:
            Term1 = 1.0# + E2onK02NuRho / 2.0#
            Term2 = (E2onK02NuRho2 / 24.0#) * (4.0# * Psi * (1.0# - 6.0# * TanFPLat2) - 3.0# * (1.0# - 16.0# * TanFPLat2) - 24.0# * TanFPLat2 / Psi)
            Term3 = E2onK02NuRho3 / 720.0#
            PointScale = Term1 + Term2 + Term3
            PointScale = PointScale * K0

        End Sub

#End Region 'Transverse Mercator Methods ----------------------------------------------------------------------------------------------------------------------------------------------------

    End Class 'TransverseMercator

    Public Class TransverseMercatorProjectionParameters
        'The set of parameters used to define a Transverse Mercator projection

        'PROPERTIES:
        '   Name
        '   Author
        '   Code
        '   LatitudeOfNaturalOrigin
        '   LongitudeOfNaturalOrigin
        '   ScaleFactorAtNaturalOrigin
        '   FalseEasting
        '   FalseNorthing


        'The name of the Transverse Mercator projection.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Transverse Mercator projection.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Transverse Mercator projection.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'The Latitude of the natural origin (in degrees).
        Private _latitudeOfNaturalOrigin As Double
        Property LatitudeOfNaturalOrigin As Double
            Get
                Return _latitudeOfNaturalOrigin
            End Get
            Set(value As Double)
                _latitudeOfNaturalOrigin = value
            End Set
        End Property

        'The Longitude of the natural origin (in degrees).
        Private _longitudeOfNaturalOrigin
        Property LongitudeOfNaturalOrigin As Double
            Get
                Return _longitudeOfNaturalOrigin
            End Get
            Set(value As Double)
                _longitudeOfNaturalOrigin = value
            End Set
        End Property

        'The scale factor at the natural origin.
        Private _scaleFactorAtNaturalOrigin As Double
        Property ScaleFactorAtNaturalOrigin As Double
            Get
                Return _scaleFactorAtNaturalOrigin
            End Get
            Set(value As Double)
                _scaleFactorAtNaturalOrigin = value
            End Set
        End Property


        Private _FalseEasting As Double
        Property FalseEasting As Double
            Get
                Return _FalseEasting
            End Get
            Set(value As Double)
                _FalseEasting = value
            End Set
        End Property

        Private _FalseNorthing As Double
        Property FalseNorthing As Double
            Get
                Return _FalseNorthing
            End Get
            Set(value As Double)
                _FalseNorthing = value
            End Set
        End Property

        'The units used to measure eastings, northings, FalseEasting and FalseNorthing
        Private _distanceUnits As String
        Property DistanceUnits As String
            Get
                Return _distanceUnits
            End Get
            Set(value As String)
                _distanceUnits = value
            End Set
        End Property

    End Class 'TransverseMercatorProjectionParameters

    Public Class GeographicCRSParameters
        'The set of Geographic CRS parameters used in the Transverse Mercator projection calculations

        'PROPERTIES:
        '   Name
        '   Author
        '   Code
        '   DatumName
        '   EllipsoidName
        '   MajorSemiAxis
        '   InverseFlattening
        '   MinorSemiAxis

        'The name of the Geographic Coordinate Reference System.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Geographic Coordinate Reference System.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the Geographic Coordinate Reference System.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'The name of the datum. The datum defines the reference ellipsoid.
        Private _datumName As String = ""
        Property DatumName As String
            Get
                Return _datumName
            End Get
            Set(value As String)
                _datumName = value
            End Set
        End Property

        'The name of the reference ellipsoid.
        Private _ellipsoidName As String = ""
        Property EllipsoidName As String
            Get
                Return _ellipsoidName
            End Get
            Set(value As String)
                _ellipsoidName = value
            End Set
        End Property

        'The semi major axis of the reference ellipsoid
        Private _semiMajorAxis As Double = Double.NaN
        Property SemiMajorAxis As Double
            Get
                Return _semiMajorAxis
            End Get
            Set(value As Double)
                _semiMajorAxis = value
            End Set
        End Property

        'The inverse flattening of the reference ellipsoid
        Private _inverseFlattening As Double = Double.NaN
        Property InverseFlattening As Double
            Get
                Return _inverseFlattening
            End Get
            Set(value As Double)
                _inverseFlattening = value
            End Set
        End Property

        'The semi minor axis of the reference ellipsoid. If this value is entered, the InverseFlattening value is calculated.
        Private _semiMinorAxis As Double = Double.NaN
        Property SemiMinorAxis As Double
            Get
                Return _semiMinorAxis
            End Get
            Set(value As Double)
                _semiMinorAxis = value
                If _semiMajorAxis = Double.NaN Then
                    'Cannot calculate inverse flattening
                Else
                    _inverseFlattening = _semiMajorAxis / (_semiMajorAxis - _semiMinorAxis)
                End If
            End Set
        End Property

    End Class 'GeographicCRSParameters

    Public Class Projection
        'Projection data.
        'The class stores the parameters of each projection.
        'Different projection types use a different set of parameters.
        'The Method parameters are used to identify the detailed parameters in the Projection Method list.
        'THe ParameterValue list contains the parameter values and units for each projection.


        Public AliasName As New List(Of String) 'Used to store alias names for the Projection.
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use identifier for this Projection. 'This is used to identify the full set of parameters in the Area Of Use list.
        Public Method As New NameAuthorCodeInfo 'Used to store the Method identifier for this Projection. 'This is used to identify the full list of parameters in the Methods list.
        'Public Parameter As New List(Of clsParameter) 'Used to store the set of projection parameters.
        Public ParameterValue As New List(Of ValueSummary) 'Used to store the set of projection parameter values.

        'The name of the projection.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the projection.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the projection.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Projection Summary has been selected. This is used to process or save a subset of Projection Summaries in a list.
        Private _selected As Boolean = True
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Comments on the projection.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property


        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property


        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Projection alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Projection alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'Projection

    Public Class CoordOpMethod
        'Coordinate Operation Method.
        'These methods include Coordinate Projections and Datum Transformations.

        Public AliasName As New List(Of String) 'Used to store alias names for the Coordinate Operation Method
        Public Parameter As New List(Of ParameterSummary) 'Used to store the list of parameters used for the Operation Method. The parameter summary does not include the parameter value and units. Different values and units are used for different Operation Methods.

        'The name of the Operation Method.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the Operation Method.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the Author to the Operation Method.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Coordinate Operation Method has been selected. This is used to process or save a subset of COMs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'If True, the parameters can be used in reverse.
        Private _reverseOp As Boolean
        Property ReverseOp As Boolean
            Get
                Return _reverseOp
            End Get
            Set(value As Boolean)
                _reverseOp = value
            End Set
        End Property

        'The formulas associated with this method or algorithm.
        Private _formula As String = ""
        Property Formula As String
            Get
                Return _formula
            End Get
            Set(value As String)
                _formula = value
            End Set
        End Property

        'Worked example of this operation method.
        Private _example As String = ""
        Property Example As String
            Get
                Return _example
            End Get
            Set(value As String)
                _example = value
            End Set
        End Property


        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified Operation Method alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified Operation Method alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'CoordOpMethod

    Public Class Transformation
        'Coordinate Transformation data.
        'The class stores the parameters of each Coordinate Transformation.
        'Different projection types use a different set of parameters.
        'The Method parameters are used to identify the detailed parameters in the Projection Method list.
        'THe ParameterValue list contains the parameter values and units for each projection.


        Public AliasName As New List(Of String) 'Used to store alias names for the Projection.
        Public SourceCRS As New NameAuthorCodeInfo
        Public TargetCRS As New NameAuthorCodeInfo
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use identifier for this Projection. 'This is used to identify the full set of parameters in the Area Of Use list.
        Public Method As New NameAuthorCodeInfo 'Used to store the Method identifier for this Projection. 'This is used to identify the full list of parameters in the Methods list.
        Public SourceCoordDiffUnit As New NameAuthorCodeInfo 'Unit of measure of the input or source coordinate differences in a polynomial operation.  Often different from the UOM of the coordinate reference system.
        Public TargetCoordDiffUnit As New NameAuthorCodeInfo 'Unit of measure of the output or target coordinate differences in a polynomial operation.  Often different from the UOM of the coordinate reference system.
        Public ParameterValue As New List(Of ValueSummary) 'Used to store the set of transformation parameter values.


        'The name of the transformation.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of the transformation.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The integer code number assigned by the author to the transformation.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this transformation Summary has been selected. This is used to process or save a subset of transformation Summaries in a list.
        Private _selected As Boolean = True
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'The version of this coordinate transforamtion
        'The version of the  transformation between these source and target coordinate reference systems.  Not required for conversions. For  transformations (datum or concatenated) may act as a secondary triple key with source and target coordinate ref systems.
        Private _version As String = ""
        Property Version As String
            Get
                Return _version
            End Get
            Set(value As String)
                _version = value
            End Set
        End Property

        'The counter for the transformation between this source and this target coordinate systems.  Not required for conversions.  In EPSG prior to v5.0 acted as the version identifier.  Retained only for purposes of backward compatibility.
        Private _variantNo As Integer
        Property VariantNo As Integer
            Get
                Return _variantNo
            End Get
            Set(value As Integer)
                _variantNo = value
            End Set
        End Property

        'The accuracy of the transformation in metres.
        Private _accuracy As Single = 0
        Property Accuracy As Single
            Get
                Return _accuracy
            End Get
            Set(value As Single)
                _accuracy = value
            End Set
        End Property


        'Comments on the transformation.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property


        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'This property is a copy of the property in the corresponding coordinate operation method.
        'If True, the parameters can be used in reverse.
        Private _reverseOp As Boolean
        Property ReverseOp As Boolean
            Get
                Return _reverseOp
            End Get
            Set(value As Boolean)
                _reverseOp = value
            End Set
        End Property


        '"Yes" = data is deprecated; "No" =  data is current and valid.
        Private _deprecated As Boolean
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified transformation alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified transformation alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

    End Class 'Transformation

    Public Class ValueSummary
        'A Value summary consists of a value quantity and summary of associated units.

        'PROPERTIES:
        'Name                       The name of the property.
        'Value                      The value of the property.
        'Unit (Name, Author, Code)  The unit of the property.

        Public Unit As New UnitOfMeasureSummary 'Stores a summary of the Parameter unit of measure (Name, Author and Code)

        'The name of the property.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The value of the property.
        Private _value As Double
        Property Value As Double
            Get
                Return _value
            End Get
            Set(value As Double)
                _value = value
            End Set
        End Property

    End Class 'ValueSummary

    Public Class Value
        'A Value consists of a value quality and associated units.

        Public Unit As New UnitOfMeasure 'Stores details of the Parameter unit of measure

        'The name of the property.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The value of the property.
        Private _value As Double
        Property Value As Double
            Get
                Return _value
            End Get
            Set(value As Double)
                _value = value
            End Set
        End Property

    End Class 'Value

    Public Class ParameterSummary
        'The Parameter Summary class stores coordinate operation parameters - excluding the parameter value and units.

        'Parameter.Name         (from Coordinate_Operation Parameter table)
        'Parameter.Description  (from Coordinate_Operation Parameter table)
        'Parameter.Order        (from Coordinate_Operation Parameter Usage table)
        'Parameter.SignReversal (from Coordinate_Operation Parameter Usage table)

        'NOTE: the fields Author, Code, Deprecated are not included in this class.
        'These would only be required if a separate list of parameters was used to store the data.
        'Parameter information is included in each Coordinate Operation Method.

        Public AliasName As New List(Of String) 'Used to store alias names for the Parameter.

        'The name of the parameter.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'A description of the property.
        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        'Sequence number indicating the order in which the parameters are shown. (From EPSG table Coordinate_Operation Parameter Usage.)
        Private _order As Integer = 0
        Property Order As Integer
            Get
                Return _order
            End Get
            Set(value As Integer)
                _order = value
            End Set
        End Property

        'Indicates if the sign of the parameter should be reversed in the reverse operation. (From EPSG table Coordinate_Operation Parameter Usage.) Only valid if ReversOp in OperationMethod is True.
        Private _signReversal As Boolean
        Property SignReversal As Boolean
            Get
                Return _signReversal
            End Get
            Set(value As Boolean)
                _signReversal = value
            End Set
        End Property

    End Class 'ParameterSummary

    Public Class Parameter
        'The Parameter class stores coordinate operation parameters

        'Parameter.Name         (from Coordinate_Operation Parameter table)
        'Parameter.Description  (from Coordinate_Operation Parameter table)
        'Parameter.Value        (from Coordinate_Operation Parameter Value table)
        'Parameter.Value.Unit.  (from Unit of Measure table)
        'Parameter.Order        (from Coordinate_Operation Parameter Usage table)
        'Parameter.SignReversal (from Coordinate_Operation Parameter Usage table)

        Public AliasName As New List(Of String) 'Used to store alias names for the Parameter.
        Public Unit As New UnitOfMeasure 'Stores details of the Parameter unit of measure

        'The name of the parameter.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'A description of the property.
        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        'The value of the property.
        Private _value As Double
        Property Value As Double
            Get
                Return _value
            End Get
            Set(value As Double)
                _value = value
            End Set
        End Property

        'Sequence number indicating the order in which the parameters are shown. (From EPSG table Coordinate_Operation Parameter Usage.)
        Private _order As Integer = 0
        Property Order As Integer
            Get
                Return _order
            End Get
            Set(value As Integer)
                _order = value
            End Set
        End Property

        'Indicates if the sign of the parameter should be reversed in the reverse operation. (From EPSG table Coordinate_Operation Parameter Usage.) Only valid if ReversOp in OperationMethod is True.
        Private _signReversal As Boolean
        Property SignReversal As Boolean
            Get
                Return _signReversal
            End Get
            Set(value As Boolean)
                _signReversal = value
            End Set
        End Property

    End Class 'Parameter


    Public Class DatumSummary
        'The DatumSummary class is used to store summary parameters of different datums.
        'Different datum types use a different set of parameters.
        'Each datum type has its own class to store the corresponding set of parameters.

        'PROPERTIES:
        'Name
        'Author
        'Code
        'Type (Geodetic, Vertical, Engineering, Image)
        'OriginDescription
        'Epoch
        'Scope
        'Comments
        'Deprecated

        '(Ellipsoid and Prime Meridian are not included in the summary parameters since these are used only for Geodetic datums.)

        'AreaOfUse

        'METHODS:
        'AliasName - List of alias names for the datum.
        'AddAlias  - Add a new name to the list of alias names.
        'Clear

        'Public Enum EnumDatumType
        Public Enum DatumTypes
            Geodetic
            Vertical
            Engineering
            Image
            Unknown
        End Enum

        Public AliasName As New List(Of String) 'used to store alias names for the datum 

        'Public AreaOfUse As New clsAreaOfUse

        'The name of the datum
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this datum summary.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to these parameters. Used in the EPSG database as the primary key in the Datum table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Datum has been selected. This is used to process or save a subset of Datums in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'The type of datum (Geodetic, Vertical, Engineering, Image)
        Private _type As DatumTypes = DatumTypes.Unknown
        Property Type As DatumTypes
            Get
                Return _type
            End Get
            Set(value As DatumTypes)
                _type = value
            End Set
        End Property

        'Origin description - A description of the anchor point, origin or datum definition.
        Private _originDescription As String = ""
        Property OriginDescription As String
            Get
                Return _originDescription
            End Get
            Set(value As String)
                _originDescription = value
            End Set
        End Property

        'The year in which the datum was realized.
        Private _epoch As String = ""
        Property Epoch As String
            Get
                Return _epoch
            End Get
            Set(value As String)
                _epoch = value
            End Set
        End Property

        'Scope of the datum
        Private _scope As String = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments about the datum.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified datum alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified ellipsoid alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Overridable Sub Clear()
            'Sets all the properties to blank or default values.
            'This is set to overridable so that derived classes can use a modified Clear Method to clear additional properties.

            AliasName.Clear()
            _name = ""
            _author = ""
            _code = 0
            _selected = False
            _type = DatumTypes.Unknown
            _originDescription = ""
            _epoch = ""
            _scope = ""
            _comments = ""
            _deprecated = True

        End Sub

    End Class 'DatumSummary

    Public Class DatumSummaryWithArea
        'The DatumSummary class is used to store summary parameters of different datums.
        'This DatumSummaryWithArea class adds the Area of Use property.

        Inherits DatumSummary

        Public AreaOfUse As New AreaOfUse

        Public Overrides Sub Clear()
            'This Clear method clears the AreaOfUse properties in addition to the DatumSummary properties.
            MyBase.Clear()
            AreaOfUse.Clear()
        End Sub

    End Class 'DatumSummaryWithArea

    Public Class GeodeticDatum_Old
        'The Geodetic Datum class is used to store the parameters of different geodetic datums. These are used to define point locations.

        Public Ellipsoid As New Ellipsoid
        Public PrimeMeridian As New PrimeMeridian
        Public AreaOfUse As New AreaOfUse

        'PROPERTIES:
        'Name
        'Author
        'Code
        'Type  (Geodetic, Vertical, Engineering, Image) (Only Geodetic stored using this class)
        'OriginDescription
        'Epoch
        'Scope
        'Comments
        'Deprecated

        'Ellipsoid
        'PrimeMeridian
        'AreaOfUse

        'METHODS:
        'AliasName - List of alias names for the datum.
        'AddAlias  - Add a new name to the list of alias names.
        'Clear


        'Private Shared AliasName As New List(Of String) 'used to store alias names for the datum 
        Public AliasName As New List(Of String) 'used to store alias names for the datum 

        'The name of the datum
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of datum parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to these parameters. Used in the EPSG database as the primary key in the Datum table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Datum has been selected. This is used to process or save a subset of Datums in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Origin description - A description of the anchor point, origin or datum definition.
        Private _originDescription As String = ""
        Property OriginDescription As String
            Get
                Return _originDescription
            End Get
            Set(value As String)
                _originDescription = value
            End Set
        End Property

        'The year in which the datum was realized.
        Private _epoch As String = ""
        Property Epoch As String
            Get
                Return _epoch
            End Get
            Set(value As String)
                _epoch = value
            End Set
        End Property

        'Scope of the datum
        Private _scope = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments about the datum.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified datum alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified ellipsoid alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Sub Clear()
            'Sets all the properties to blank or default values

            AliasName.Clear()
            _name = ""
            _author = ""
            _code = 0
            _selected = False
            _originDescription = ""
            _epoch = ""
            _scope = ""
            _comments = ""
            _deprecated = True

        End Sub

    End Class 'GeodeticDatum

    Public Class GeodeticDatum
        'The Geodetic Datum class is used to store the parameters of different geodetic datums. These are used to define point locations.

        'The old version of the GeodeticDatum class stored the complete Ellipsoid, PrimeMeridian and AreaOfUse data in each record.
        'The new version stores only the Name, Author and Code number used to identify the corresponding records in the Ellipsoid, PrimeMeridian and AreaOfuse lists.
        'Public Ellipsoid As New Ellipsoid
        'Public PrimeMeridian As New PrimeMeridian
        'Public AreaOfUse As New AreaOfUse

        Public Ellipsoid As New NameAuthorCodeInfo
        Public PrimeMeridian As New NameAuthorCodeInfo
        Public Area As New NameAuthorCodeInfo 'Used to store the area of use identifier for this Datum. 'This is used to identify the full set of parameters in the Area Of Use list.

        'PROPERTIES:
        'Name
        'Author
        'Code
        'Type  (Geodetic, Vertical, Engineering, Image) (Only Geodetic stored using this class)
        'OriginDescription
        'Epoch
        'Scope
        'Comments
        'Deprecated

        'Ellipsoid
        'PrimeMeridian
        'AreaOfUse

        'METHODS:
        'AliasName - List of alias names for the datum.
        'AddAlias  - Add a new name to the list of alias names.
        'Clear


        'Private Shared AliasName As New List(Of String) 'used to store alias names for the datum 
        Public AliasName As New List(Of String) 'used to store alias names for the datum 

        'The name of the datum
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of datum parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to these parameters. Used in the EPSG database as the primary key in the Datum table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Datum has been selected. This is used to process or save a subset of Datums in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Origin description - A description of the anchor point, origin or datum definition.
        Private _originDescription As String = ""
        Property OriginDescription As String
            Get
                Return _originDescription
            End Get
            Set(value As String)
                _originDescription = value
            End Set
        End Property

        'The year in which the datum was realized.
        Private _epoch As String = ""
        Property Epoch As String
            Get
                Return _epoch
            End Get
            Set(value As String)
                _epoch = value
            End Set
        End Property

        'Scope of the datum
        Private _scope = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments about the datum.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified datum alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified ellipsoid alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Sub Clear()
            'Sets all the properties to blank or default values

            AliasName.Clear()
            _name = ""
            _author = ""
            _code = 0
            _selected = False
            _originDescription = ""
            _epoch = ""
            _scope = ""
            _comments = ""
            _deprecated = True

        End Sub

    End Class 'GeodeticDatum

    Public Class VerticalDatum
        'The Vertical Datum class is used to store the parameters of different vertical datums. These are used to define point locations.

        Public Area As New NameAuthorCodeInfo 'Used to store the area of use identifier for this Datum. 'This is used to identify the full set of parameters in the Area Of Use list.

        Public AliasName As New List(Of String) 'used to store alias names for the datum 

        'The name of the datum
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of datum parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to these parameters. Used in the EPSG database as the primary key in the Datum table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Datum has been selected. This is used to process or save a subset of Datums in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Origin description - A description of the anchor point, origin or datum definition.
        Private _originDescription As String = ""
        Property OriginDescription As String
            Get
                Return _originDescription
            End Get
            Set(value As String)
                _originDescription = value
            End Set
        End Property

        'The year in which the datum was realized.
        Private _epoch As String = ""
        Property Epoch As String
            Get
                Return _epoch
            End Get
            Set(value As String)
                _epoch = value
            End Set
        End Property

        'Scope of the datum
        Private _scope = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments about the datum.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified datum alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified datum alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Sub Clear()
            'Sets all the properties to blank or default values

            AliasName.Clear()
            _name = ""
            _author = ""
            _code = 0
            _selected = False
            _originDescription = ""
            _epoch = ""
            _scope = ""
            _comments = ""
            _deprecated = True

        End Sub

    End Class 'VerticalDatum 

    Public Class EngineeringDatum '--------------------------------------------------------------------
        'The EngineeringDatum class is used to store the parameters of different engineering datums. These are used to define point locations.

        Public Area As New NameAuthorCodeInfo 'Used to store the area of use identifier for this Datum. 'This is used to identify the full set of parameters in the Area Of Use list.

        Public AliasName As New List(Of String) 'used to store alias names for the datum 

        'The name of the datum
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of datum parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to these parameters. Used in the EPSG database as the primary key in the Datum table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Datum has been selected. This is used to process or save a subset of Datums in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Origin description - A description of the anchor point, origin or datum definition.
        Private _originDescription As String = ""
        Property OriginDescription As String
            Get
                Return _originDescription
            End Get
            Set(value As String)
                _originDescription = value
            End Set
        End Property

        'The year in which the datum was realized.
        Private _epoch As String = ""
        Property Epoch As String
            Get
                Return _epoch
            End Get
            Set(value As String)
                _epoch = value
            End Set
        End Property

        'Scope of the datum
        Private _scope = ""
        Property Scope As String
            Get
                Return _scope
            End Get
            Set(value As String)
                _scope = value
            End Set
        End Property

        'Comments about the datum.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'If True, data is deprecated. If False, data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified datum alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is already on the alias list.
            Else
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified datum alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Sub Clear()
            'Sets all the properties to blank or default values

            AliasName.Clear()
            _name = ""
            _author = ""
            _code = 0
            _selected = False
            _originDescription = ""
            _epoch = ""
            _scope = ""
            _comments = ""
            _deprecated = True

        End Sub
    End Class 'EngineeringDatum ---------------------------------------------------------------------

    Public Class Ellipsoid
        'The ellipsoid class is used to store ellispod parameters.

        'PROPERTIES:
        'Name
        'EpsgCode
        'Comments
        'EllipsoidParameters (SemiMajorAxis_InverseFlattening or SemiMajorAxis_SemiMinorAxis
        'SemiMajorAxis
        'InverseFlattening
        'SemiMinorAxis

        'METHODS:
        'AddAlias  - Add a new name to the list of alias names.
        'RemoveAlias - Remove a name from the list of alias names.
        'Clear - Sets all the properties to blank or default values.

        Public Unit As New UnitOfMeasure 'Stores details of the Axis unit of measure

        'Ellipsoids are defined useing either SemiMajorAxis and InverseFlattening or SemiMajorAxis and SemiMinorAxis
        Public Enum DefiningParameters
            SemiMajorAxis_InverseFlattening
            SemiMajorAxis_SemiMinorAxis
            Unknown
        End Enum

        Public AliasName As New List(Of String) 'used to store alias names for the datum 

        'The name of the ellipsoid.
        Private _name As String = "" '(80 characters max)
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of ellipsoid parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to these parameters. Used in the EPSG database as the primary key in the Ellipsoid table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this Ellipsoid has been selected. This is used to process or save a subset of Ellipsoids in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Comments about the ellipsoid.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'False if data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property


        'Ellipsoids are defined useing either SemiMajorAxis and InverseFlattening or SemiMajorAxis and SemiMinorAxis.
        Private _ellipsoidParameters As DefiningParameters = DefiningParameters.Unknown
        Property EllipsoidParameters As DefiningParameters
            Get
                Return _ellipsoidParameters
            End Get
            Set(value As DefiningParameters)
                _ellipsoidParameters = value
            End Set
        End Property

        'The value of the semi major axis.
        Private _semiMajorAxis As Double = Double.NaN
        Property SemiMajorAxis As Double
            Get
                Return _semiMajorAxis
            End Get
            Set(value As Double)
                _semiMajorAxis = value
            End Set
        End Property

        'The value of the inverse flattening.
        Private _inverseFlattening As Double = Double.NaN
        Property InverseFlattening As Double
            Get
                Return _inverseFlattening
            End Get
            Set(value As Double)
                _inverseFlattening = value
            End Set
        End Property

        'The value of the semi minor axis.
        Private _semiMinorAxis As Double = Double.NaN
        Property SemiMinorAxis As Double
            Get
                Return _semiMinorAxis
            End Get
            Set(value As Double)
                _semiMinorAxis = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified ellipsoid alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified ellipsoid alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Sub Clear()
            'Sets all the properties to blank or default values.
            AliasName.Clear()
            _name = ""
            _author = ""
            _code = 0
            _selected = False
            _comments = ""
            _deprecated = True
            _ellipsoidParameters = DefiningParameters.Unknown
            _semiMajorAxis = Double.NaN
            _inverseFlattening = Double.NaN
            _semiMinorAxis = Double.NaN

        End Sub

    End Class 'Ellipsoid

    Public Class PrimeMeridian
        'The PrimeMeridian class stores Prime Meridian parmeters, which are used to define a datum.

        'NOTE: LongitudeFromGreenwich is stored using decimal degrees units of measure.

        'PROPERTIES:
        'Name
        'EpsgCode
        'Comments
        'LongitudeUOM (DegMinSec, Grad)
        'LongitudeFromGreenwich

        'AliasName - List of alias names for the Area Of Use.

        'METHODS:
        'AddAlias  - Add a new name to the list of alias names.
        'RemoveAlias - Remove a name from the list of alias names.
        'Clear - Sets all the properties to blank or default values.

        'Public Enum EnumLongitudeUnits
        Public Enum LongitudeUnits
            'DegMinSec
            Degree
            Gradian
            Sexagesimal_DMS
            Unknown
        End Enum

        '9102 - degree
        '9105 - grad
        '9110 - sexagesimal DMS - (Base 60)

        Public AliasName As New List(Of String) 'Used to store alias names for the prime meridian

        'The name of the Prime Meridian
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of prime meridian parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to these parameters. Used in the EPSG database as the primary key in the Prime Meridian table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this PM has been selected. This is used to process or save a subset of PMs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Comments about the prime meridian.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'False if data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        Private _longitudeUOM As LongitudeUnits = LongitudeUnits.Degree
        Property LongitudeUOM As LongitudeUnits
            Get
                Return _longitudeUOM
            End Get
            Set(value As LongitudeUnits)
                _longitudeUOM = value
            End Set
        End Property

        'The longitude of the prime meridian in the Greenwich system. Units in radians.
        Private _longitudeFromGreenwich As Double = 0
        Property LongitudeFromGreenwich As Double
            Get
                Return _longitudeFromGreenwich
            End Get
            Set(value As Double)
                _longitudeFromGreenwich = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified prime meridian alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified prime meridian alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Sub Clear()
            'Sets all the properties to blank or default values.
            _name = ""
            _author = ""
            _code = 0
            _comments = ""
            _deprecated = True
            _longitudeUOM = LongitudeUnits.Degree
            _longitudeFromGreenwich = 0

        End Sub

    End Class 'PrimeMeridian

    Public Class UnitOfMeasureSummary
        'Unit of measure summary parameters.
        'These parameters include only Name, Author and Code.
        'These parameters are sufficent to identify the full set of parameters from a parameter list.

        'PROPERTIES:
        'Name
        'Author
        'Code

        'The name of the unit of measure.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of unit of measure parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to this Unit of Measure. Used in the EPSG database as the primary key in the Unit of Measure table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

    End Class 'UnitOfMeasureSummary

    Public Class UnitOfMeasure '--------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Unit of measure parameters for different measures used to define coordinates and coordinate conversions.
        '
        'http://www.energistics.org/asset-data-management/unit-of-measure-standard

        'Public Enum EnumUOMType
        Public Enum UOMTypes
            Scale
            Length
            Time
            Angle
            Unknown
        End Enum

        Public AliasName As New List(Of String) 'Used to store alias names for the unit of measure

        'PROPERTIES:
        'Name
        'Author
        'Code
        'Type
        'Comments
        'Deprecated
        'FactorB
        'FactorC
        'StandardUnitName

        'METHODS:
        'ConvertToStandardUnits
        'ConvertFromStandardUnits
        'AddAlias
        'RemoaveAlias
        'Clear - Sets all the properties to blank or default values.

        'The name of the unit of measure.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of unit of measure parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code assigned by the author to this Unit of Measure. Used in the EPSG database as the primary key in the Unit of Measure table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this UOM has been selected. This is used to process or save a subset of UOMs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'The unit of measure type.
        Private _type As UOMTypes = UOMTypes.Unknown
        Property Type As UOMTypes
            Get
                Return _type
            End Get
            Set(value As UOMTypes)
                _type = value
            End Set
        End Property

        'Comments on the Unit of Measure
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'False if data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        'General unit conversion factors are A, B, C and D and are used in the formula y=(A+Bx)/(C+Dx) 
        'Only factors B and C are needed here.

        'Factor B used in the conversion formula y = (B/C).x
        Private _factorB As Double
        Property FactorB As Double
            Get
                Return _factorB
            End Get
            Set(value As Double)
                _factorB = value
            End Set
        End Property

        'Factor C used in the conversion formula y = (B/C).x
        Private _factorC As Double
        Property FactorC As Double
            Get
                Return _factorC
            End Get
            Set(value As Double)
                _factorC = value
            End Set
        End Property

        'The name of the unit of measure converted using the formula y = (B/C).x
        Private _standardUnitName As String
        Property StandardUnitName As String
            Get
                Return _standardUnitName
            End Get
            Set(value As String)
                _standardUnitName = value
            End Set
        End Property

        Public Function ConvertToStandardUnits(ByVal Measure As Double) As Double
            'Convert a measure into StandardUnitName units.

            Return Measure * _factorB / _factorC

            'If _factorC = 0 Then
            '    'Divide by zero error
            'Else
            '    Return Measure * _factorB / _factorC
            'End If

        End Function

        Public Function ConvertFromStandardUnits(ByVal Measure As Double) As Double
            'Convert a measure from Standard Units to this unit of measure
            Return Measure * _factorC / _factorB
        End Function

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified UOM alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified UOM alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Sub Clear()
            'Sets all the properties to blank or default values.
            _name = ""
            _author = ""
            _code = 0
            _selected = False
            _type = UOMTypes.Unknown
            _comments = ""
            _deprecated = True

        End Sub

    End Class 'UnitOfMeasure -----------------------------------------------------------------------------------------------------------------------------------------------------------------

    'CHECK LAT AND LONG UNITS!!! (WGS84???)
    Public Class AreaOfUse
        'The Area Of Use class is used to store Area Of Use parameters.

        'PROPERTIES:
        'Name
        'Code
        'Comments
        'Deprecated
        'AreaOfUse
        'SouthLatitude
        'NorthLatitude
        'WestLongitude
        'EastLongitude
        'IsoA2Code
        'IsoA3Code
        'IsoNCode

        'AliasName - List of alias names for the Area Of Use.

        'METHODS:
        'AddAlias  - Add a new name to the list of alias names.
        'RemoveAlias - Remove a name from the list of alias names.
        'Clear - Sets all the properties to blank or default values.

        Public AliasName As New List(Of String) 'used to store alias names for the Area Of Use. 

        'The name of the Area Of Use.
        Private _name As String = ""
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The author of this set of area of use parameters.
        Private _author As String = ""
        Property Author As String
            Get
                Return _author
            End Get
            Set(value As String)
                _author = value
            End Set
        End Property

        'The unique code used in the EPSG database as the primary key in the Area table.
        Private _code As Integer = 0
        Property Code As Integer
            Get
                Return _code
            End Get
            Set(value As Integer)
                _code = value
            End Set
        End Property

        'If True, this AOU has been selected. This is used to process or save a subset of AOUs in a list.
        Private _selected As Boolean = False
        Property Selected As Boolean
            Get
                Return _selected
            End Get
            Set(value As Boolean)
                _selected = value
            End Set
        End Property

        'Comments about the Area Of Use.
        Private _comments As String = ""
        Property Comments As String
            Get
                Return _comments
            End Get
            Set(value As String)
                _comments = value
            End Set
        End Property

        'False if data is current and valid.
        Private _deprecated As Boolean = True
        Property Deprecated As Boolean
            Get
                Return _deprecated
            End Get
            Set(value As Boolean)
                _deprecated = value
            End Set
        End Property

        'Description of the area of use.
        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _southLatitude As Double = Double.NaN
        Property SouthLatitude As Double
            Get
                Return _southLatitude
            End Get
            Set(value As Double)
                _southLatitude = value
            End Set
        End Property

        Private _northLatitude As Double = Double.NaN
        Property NorthLatitude As Double
            Get
                Return _northLatitude
            End Get
            Set(value As Double)
                _northLatitude = value
            End Set
        End Property

        Private _westLongitude As Double = Double.NaN
        Property WestLongitude As Double
            Get
                Return _westLongitude
            End Get
            Set(value As Double)
                _westLongitude = value
            End Set
        End Property

        Private _eastLongitude As Double = Double.NaN
        Property EastLongitude As Double
            Get
                Return _eastLongitude
            End Get
            Set(value As Double)
                _eastLongitude = value
            End Set
        End Property

        Private _isoA2Code As String = ""
        Property IsoA2Code As String 'ISO 3166 2-digit alpha country code
            Get
                Return _isoA2Code
            End Get
            Set(value As String)
                _isoA2Code = value
            End Set
        End Property

        Private _isoA3Code As String = ""
        Property IsoA3Code As String 'ISO 3166 3-digit alpha country code
            Get
                Return _isoA3Code
            End Get
            Set(value As String)
                _isoA3Code = value
            End Set
        End Property

        Private _isoNCode As Integer = 0
        Property IsoNCode As Integer 'ISO 3166 3-digit numeric country code
            Get
                Return _isoNCode
            End Get
            Set(value As Integer)
                _isoNCode = value
            End Set
        End Property

        Public Sub AddAlias(ByVal Name As String)
            'Add the specified AOU alias name to the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
            Else
                'The name is not on the alias list.
                AliasName.Add(Name)
            End If
        End Sub

        Public Sub RemoveAlias(ByVal Name As String)
            'Remove the specified AOU alias name from the list.

            'Check if the name is already on the list:
            If AliasName.Contains(Name) Then
                'The name is on the alias list.
                AliasName.Remove(Name)
            Else
                'The name is not on the alias list.
            End If
        End Sub

        Public Sub Clear()
            'Sets all the properties to blank or default values.
            _name = ""
            _author = ""
            _code = 0
            _selected = False
            _comments = ""
            _deprecated = True
            _description = ""
            _southLatitude = Double.NaN
            _northLatitude = Double.NaN
            _westLongitude = Double.NaN
            _eastLongitude = Double.NaN
            _isoA2Code = ""
            _isoA3Code = ""
            _isoNCode = 0

        End Sub

    End Class 'AreaOfUse

    Public Class SevenParamTransformation
        'This class is used to store the parameters used for a datum transformation.
        'Seven parameter similarity transformation parameters.
        'AKA 3 Dimensional Similarity Transformation.

        'The name of the datum transformation.
        Private _name As String
        Property Name As String
            Get
                Return _name
            End Get
            Set(value As String)
                _name = value
            End Set
        End Property

        'The datum that the coordinates are being transformaed from.
        Private _fromDatum As String
        Property FromDatum As String
            Get
                Return _fromDatum
            End Get
            Set(value As String)
                _fromDatum = value
            End Set
        End Property

        'The datum that the coordinates are being transformed to.
        Private _toDatum As String
        Property ToDatum As String
            Get
                Return _toDatum
            End Get
            Set(value As String)
                _toDatum = value
            End Set
        End Property

        'The X-axis translation in metres.
        Private _dx As Double
        Property Dx As Double
            Get
                Return _dx
            End Get
            Set(value As Double)
                _dx = value
            End Set
        End Property

        'The Y-axis translation in metres.
        Private _dy As Double
        Property Dy As Double
            Get
                Return _dy
            End Get
            Set(value As Double)
                _dy = value
            End Set
        End Property

        'The Z-axis translation in metres.
        Private _dz As Double
        Property Dz As Double
            Get
                Return _dz
            End Get
            Set(value As Double)
                _dz = value
            End Set
        End Property

        'The X-axis rotation in arc-seconds.
        Private _rx As Double
        Property Rx As Double
            Get
                Return _rx
            End Get
            Set(value As Double)
                _rx = value
            End Set
        End Property

        'The Y-axis rotation in arc-seconds.
        Private _ry As Double
        Property Ry As Double
            Get
                Return _ry
            End Get
            Set(value As Double)
                _ry = value
            End Set
        End Property

        'The Z-axis rotation in arc-seconds.
        Private _rz As Double
        Property Rz As Double
            Get
                Return _rz
            End Get
            Set(value As Double)
                _rz = value
            End Set
        End Property

        'The scale difference in parts per million.
        Private _sc As Double
        Property Sc As Double
            Get
                Return _sc
            End Get
            Set(value As Double)
                _sc = value
            End Set
        End Property

    End Class 'SevenParamTransformation

    Public Class AngleDegMinSec
        'Stores an angle expressed as degress, minutes and seconds.
        'Converts between DecimalDegrees and DegreesMinutesSeconds:
        'Function DegMinSecToDecimalDegrees() As Double

        Public Enum Sign
            Positive
            Negative
        End Enum

        Private _degMinSecSign As Sign = Sign.Positive
        Property DegMinSecSign As Sign
            Get
                Return _degMinSecSign
            End Get
            Set(value As Sign)
                _degMinSecSign = value
            End Set
        End Property

        Private _degrees As Integer
        Property Degrees As Integer
            Get
                Return _degrees
            End Get
            Set(value As Integer)
                _degrees = value
            End Set
        End Property

        Private _minutes As Integer
        Property Minutes As Integer
            Get
                Return _minutes
            End Get
            Set(value As Integer)
                _minutes = value
            End Set
        End Property

        Private _seconds As Decimal
        Property Seconds As Decimal
            Get
                Return _seconds
            End Get
            Set(value As Decimal)
                _seconds = value
            End Set
        End Property

        Private _secondsDecimalPlaces As Integer = 4
        Property SecondsDecimalPlaces As Integer 'The number of decimal places used to display the seconds value.
            Get
                Return _secondsDecimalPlaces
            End Get
            Set(value As Integer)
                _secondsDecimalPlaces = value
            End Set
        End Property

        Public Sub DecimalDegreesToDegMinSec_Old(ByVal DecimalDegrees As Double)
            'Set the Degrees, Minutes and Seconds values to the same angle as DecimalDegrees

            'In this version of the subroutine negative angles are represented using negative Degrees, Minutes and Seconds.
            'The updated version uses the DegMinSec sign to indicate the sign of the Deg Min Sec angle.

            'Int(99.8) = 99
            'Fix(99.8) = 99
            'Int(-99.8) = -100
            'Fix(-99.8) = -99

            'Degrees = Int(DecimalDegrees)
            Degrees = Fix(DecimalDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (DecimalDegrees - Degrees) * 60
            Minutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, SecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes >= 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes <= -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If

            'Seconds = (DecimalMinutes - Minutes) * 60
            Seconds = RoundedSeconds
        End Sub

        Public Sub DecimalDegreesToDegMinSec(ByVal DecimalDegrees As Double)
            'Set the Degrees, Minutes and Seconds values to the same angle as DecimalDegrees

            'If DecimalDegrees is negative, DegMinSecSign is set to Negative.

            Dim AbsDecimalDegrees As Double
            If DecimalDegrees < 0 Then
                AbsDecimalDegrees = Math.Abs(DecimalDegrees)
                DegMinSecSign = Sign.Negative
            Else
                AbsDecimalDegrees = DecimalDegrees
                DegMinSecSign = Sign.Positive
            End If

            Degrees = Fix(AbsDecimalDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecimalDegrees - Degrees) * 60
            Minutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, SecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes >= 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes <= -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If

            Seconds = RoundedSeconds

        End Sub

        Public Function DegMinSecToDecimalDegrees_Old() As Double
            'Convert the Degrees, Minutes and Seconds values to a Decimal Degrees value.
            Return Degrees + Minutes / 60 + Seconds / 3600
        End Function

        Public Function DegMinSecToDecimalDegrees() As Double
            'Convert the Degrees, Minutes and Seconds values to a Decimal Degrees value.
            If DegMinSecSign = Sign.Negative Then
                Return (Degrees + Minutes / 60 + Seconds / 3600) * -1
            Else
                Return Degrees + Minutes / 60 + Seconds / 3600
            End If
        End Function

    End Class 'AngleDegMinSec

    Public Class AngleConvert
        'Converts angles between turns, decimal degrees, sexagesimal degrees, radians and grads (or gradians)
        'UPDATE: Now includes angles in degrees, minutes & seconds.

#Region "Angle properties"

        Private _radians As Double
        Property Radians As Double
            Get
                Return _radians
            End Get
            Set(value As Double)
                _radians = value
            End Set
        End Property

        Private _turns As Double
        Property Turns As Double
            Get
                Return _turns
            End Get
            Set(value As Double)
                _turns = value
            End Set
        End Property

        Private _decimalDegrees As Double
        Property DecimalDegrees As Double
            Get
                Return _decimalDegrees
            End Get
            Set(value As Double)
                _decimalDegrees = value
            End Set
        End Property

        Private _sexagesimalDegrees As Double
        Property SexagesimalDegrees As Double
            Get
                Return _sexagesimalDegrees
            End Get
            Set(value As Double)
                _sexagesimalDegrees = value
            End Set
        End Property

        Private _gradians As Double
        Property Gradians As Double
            Get
                Return _gradians
            End Get
            Set(value As Double)
                _gradians = value
            End Set
        End Property

#End Region 'Angle properties: decimal degrees, sexagesimal degrees, radians, gradians & turns.


#Region "Degrees, Minutes & Seconds properties"

        Public Enum Sign
            Positive
            Negative
        End Enum

        Private _dmsSign As Sign = Sign.Positive
        Property DmsSign As Sign
            Get
                Return _dmsSign
            End Get
            Set(value As Sign)
                _dmsSign = value
            End Set
        End Property

        Private _dmsDegrees As Integer
        Property DmsDegrees As Integer
            Get
                Return _dmsDegrees
            End Get
            Set(value As Integer)
                _dmsDegrees = value
            End Set
        End Property

        Private _dmsMinutes As Integer
        Property DmsMinutes As Integer
            Get
                Return _dmsMinutes
            End Get
            Set(value As Integer)
                _dmsMinutes = value
            End Set
        End Property


        Private _dmsSeconds As Decimal
        Property DmsSeconds As Decimal
            Get
                Return _dmsSeconds
            End Get
            Set(value As Decimal)
                _dmsSeconds = value
            End Set
        End Property

        Private _dmsSecondsDecimalPlaces As Integer = 4
        Property DmsSecondsDecimalPlaces As Integer 'The number of decimal places used to display the seconds value.
            Get
                Return _dmsSecondsDecimalPlaces
            End Get
            Set(value As Integer)
                _dmsSecondsDecimalPlaces = value
            End Set
        End Property

#End Region 'Degrees, Minutes & Seconds properties

#Region "Convert Radians"

        'Radians to: Decimal degrees, Sexagesimal degrees, Turns, Gradians -----------------------------------------------------
        Public Sub ConvertRadianToDecimalDegree()
            DecimalDegrees = Radians * 360 / 2 / System.Math.PI
        End Sub

        Public Sub ConvertRadianToSexagesimalDegree()
            Dim Degrees As Integer = Fix(Radians * 360 / 2 / System.Math.PI)
            Dim DecimalMinutes As Double = (Radians * 360 / 2 / System.Math.PI - Degrees) * 60
            Dim Minutes As Integer = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal = System.Decimal.Round(RawSeconds, 6) 'Round to 6 decimal places
            'Check if RoundedSeconds >= 60
            If RoundedSeconds >= 60 Then 'Rounding has increased the number of seconds past 60.
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes > 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then 'Rounding has decreased the number of (negative) seconds below -60
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes < -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If
            SexagesimalDegrees = Degrees + Minutes / 100 + RoundedSeconds / 10000
        End Sub

        Public Sub ConvertRadianToTurn()
            Turns = Radians / 2 / System.Math.PI
        End Sub

        Public Sub ConvertRadianToGradian()
            Gradians = Radians * 400 / 2 / System.Math.PI
        End Sub

        Public Sub ConvertRadianToDegMinSec()
            'Convert Radian to Degrees, Minutes & Seconds
            Dim DecDegrees As Double
            DecDegrees = Radians * 360 / 2 / System.Math.PI

            Dim AbsDecDegrees As Double
            If DecDegrees < 0 Then
                AbsDecDegrees = Math.Abs(DecDegrees)
                DmsSign = Sign.Negative
            Else
                AbsDecDegrees = DecDegrees
                DmsSign = Sign.Positive
            End If

            DmsDegrees = Fix(AbsDecDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecDegrees - DmsDegrees) * 60
            DmsMinutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - DmsMinutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, DmsSecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                DmsMinutes = DmsMinutes + 1
                RoundedSeconds = 0
                If DmsMinutes >= 60 Then
                    DmsDegrees = DmsDegrees + 1
                    DmsMinutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                DmsMinutes = DmsMinutes - 1
                RoundedSeconds = 0
                If DmsMinutes <= -60 Then
                    DmsDegrees = DmsDegrees - 1
                    DmsMinutes = 0
                End If
            End If
            DmsSeconds = RoundedSeconds
        End Sub

#End Region 'Convert Radians - Methods to convert radians to other angle units.


#Region "Convert Decimal Degrees"

        'Decimal degrees to: Radians, Sexagesimal degrees, Turns, Gradians -------------------------------------------------------
        Public Sub ConvertDecimalDegreeToRadian()
            Radians = DecimalDegrees * 2 * System.Math.PI / 360
        End Sub

        Public Sub ConvertDecimalDegreeToSexagesimalDegree()
            Dim Degrees As Integer = Fix(DecimalDegrees)
            Dim DecimalMinutes As Double = (DecimalDegrees - Degrees) * 60
            Dim Minutes As Integer = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal = System.Decimal.Round(RawSeconds, 6) 'Round to 6 decimal places
            'Check if RoundedSeconds >= 60
            If RoundedSeconds >= 60 Then 'Rounding has increased the number of seconds past 60.
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes > 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then 'Rounding has decreased the number of (negative) seconds below -60
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes < -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If
            SexagesimalDegrees = Degrees + Minutes / 100 + RoundedSeconds / 10000
        End Sub

        Public Sub ConvertDecimalDegreeToTurn()
            Turns = DecimalDegrees / 360
        End Sub

        Public Sub ConvertDecimalDegreeToGradian()
            Gradians = DecimalDegrees * 400 / 360
        End Sub

        Public Sub ConvertDecimalDegreeToDegMinSec()
            'Convert DecimalDegree to Degrees, Minutes & Seconds
            Dim AbsDecimalDegrees As Double
            If DecimalDegrees < 0 Then
                AbsDecimalDegrees = Math.Abs(DecimalDegrees)
                DmsSign = Sign.Negative
            Else
                AbsDecimalDegrees = DecimalDegrees
                DmsSign = Sign.Positive
            End If

            DmsDegrees = Fix(AbsDecimalDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecimalDegrees - DmsDegrees) * 60
            DmsMinutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - DmsMinutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, DmsSecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                DmsMinutes = DmsMinutes + 1
                RoundedSeconds = 0
                If DmsMinutes >= 60 Then
                    DmsDegrees = DmsDegrees + 1
                    DmsMinutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                DmsMinutes = DmsMinutes - 1
                RoundedSeconds = 0
                If DmsMinutes <= -60 Then
                    DmsDegrees = DmsDegrees - 1
                    DmsMinutes = 0
                End If
            End If
            DmsSeconds = RoundedSeconds
        End Sub

#End Region 'Convert Decimal Degrees - Methods to convert decimal degrees to other angle units.


#Region "Convert Sexagesimal Degrees"

        'Sexagesimal degrees to: Radians, Decimal degrees, Turns, Gradians ------------------------------------------------------

        Public Sub ConvertSexagesimalDegreeToRadian()
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            Dim Minutes As Integer = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            Dim Seconds As Double = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
            If SexagesimalStr.StartsWith("-") Then
                Radians = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * -2 * System.Math.PI / 360
            ElseIf SexagesimalStr.StartsWith("+") Then
                Radians = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * 2 * System.Math.PI / 360
            Else
                Radians = (SexagesimalStr.Substring(0, DecimalPointPos) + Minutes / 60 + Seconds / 3600) * 2 * System.Math.PI / 360
            End If
        End Sub


        Public Sub ConvertSexagesimalDegreeToDecimalDegree()
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            Dim Minutes As Integer = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            Dim Seconds As Double = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
            If SexagesimalStr.StartsWith("-") Then
                DecimalDegrees = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * -1
            ElseIf SexagesimalStr.StartsWith("+") Then
                DecimalDegrees = SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600
            Else
                DecimalDegrees = SexagesimalStr.Substring(0, DecimalPointPos) + Minutes / 60 + Seconds / 3600
            End If
        End Sub


        Public Sub ConvertSexagesimalDegreeToTurn()
            'Dim SexagesimalStr As String = Trim(Str(SexagesimalDegrees))
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            Dim Minutes As Integer = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            Dim Seconds As Double = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
            If SexagesimalStr.StartsWith("-") Then
                Turns = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * -1 / 360
            ElseIf SexagesimalStr.StartsWith("+") Then
                Turns = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) / 360
            Else
                Turns = (SexagesimalStr.Substring(0, DecimalPointPos) + Minutes / 60 + Seconds / 3600) / 360
            End If
        End Sub


        Public Sub ConvertSexagesimalDegreeToGradian()
            'Dim SexagesimalStr As String = Trim(Str(SexagesimalDegrees))
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            Dim Minutes As Integer = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            Dim Seconds As Double = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
            If SexagesimalStr.StartsWith("-") Then
                Gradians = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * -400 / 360
            ElseIf SexagesimalStr.StartsWith("+") Then
                Gradians = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * 400 / 360
            Else
                Gradians = (SexagesimalStr.Substring(0, DecimalPointPos) + Minutes / 60 + Seconds / 3600) * 400 / 360
            End If
        End Sub


        Public Sub ConvertSexagesimalDegreeToDegMinSec()
            'To avoid errors caused by rounding, Sexagesimal degrees are converted using character positions in the Sexagesimal string.
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            If SexagesimalStr.StartsWith("-") Then
                DmsSign = Sign.Negative
                DmsDegrees = SexagesimalStr.Substring(1, DecimalPointPos - 1)
            ElseIf SexagesimalStr.StartsWith("+") Then
                DmsSign = Sign.Positive
                DmsDegrees = SexagesimalStr.Substring(1, DecimalPointPos - 1)
            Else
                DmsSign = Sign.Positive
                DmsDegrees = SexagesimalStr.Substring(0, DecimalPointPos)
            End If
            DmsMinutes = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            DmsSeconds = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
        End Sub

#End Region 'Convert Sexagesimal Degrees - Methods to convert sexagesimal degrees to other angle units.


#Region "Convert Turns"

        'Turns to: Radian, Decimal degrees, Sexagesimal degrees, Gradians -------------------------------------------------------
        Public Sub ConvertTurnToRadian()
            Radians = Turns * 2 * System.Math.PI
        End Sub

        Public Sub ConvertTurnToDecimalDegree()
            DecimalDegrees = Turns * 360
        End Sub

        Public Sub ConvertTurnToSexagesimalDegree()
            Dim Degrees As Integer = Fix(Turns * 360)
            Dim DecimalMinutes As Double = (Turns * 360 - Degrees) * 60
            Dim Minutes As Integer = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal = System.Decimal.Round(RawSeconds, 6) 'Round to 6 decimal places
            'Check if RoundedSeconds >= 60
            If RoundedSeconds >= 60 Then 'Rounding has increased the number of seconds past 60.
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes > 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then 'Rounding has decreased the number of (negative) seconds below -60
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes < -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If
            SexagesimalDegrees = Degrees + Minutes / 100 + RoundedSeconds / 10000
        End Sub

        Public Sub ConvertTurnToGradian()
            Gradians = Turns * 400
        End Sub

        Public Sub ConvertTurnToDegMinSec()
            'Convert Turn to Degrees, Minutes & Seconds
            Dim DecDegrees As Double
            DecDegrees = Turns * 360

            Dim AbsDecDegrees As Double
            If DecDegrees < 0 Then
                AbsDecDegrees = Math.Abs(DecDegrees)
                DmsSign = Sign.Negative
            Else
                AbsDecDegrees = DecDegrees
                DmsSign = Sign.Positive
            End If

            DmsDegrees = Fix(AbsDecDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecDegrees - DmsDegrees) * 60
            DmsMinutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - DmsMinutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, DmsSecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                DmsMinutes = DmsMinutes + 1
                RoundedSeconds = 0
                If DmsMinutes >= 60 Then
                    DmsDegrees = DmsDegrees + 1
                    DmsMinutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                DmsMinutes = DmsMinutes - 1
                RoundedSeconds = 0
                If DmsMinutes <= -60 Then
                    DmsDegrees = DmsDegrees - 1
                    DmsMinutes = 0
                End If
            End If

            DmsSeconds = RoundedSeconds
        End Sub

#End Region 'Convert Turns - Methods to convert turns to other angle units.


#Region "Convert Gradians"

        'Gradians to: Radians, Decimal degrees, Sexagesimal degrees, Turns ------------------------------------------------------
        Public Sub ConvertGradianToRadian()
            Radians = Gradians * 2 * System.Math.PI / 400
        End Sub

        Public Sub ConvertGradianToDecimalDegree()
            DecimalDegrees = Gradians * 360 / 400
        End Sub

        Public Sub ConvertGradianToSexagesimalDegree()
            Dim Degrees As Integer = Fix(Gradians * 360 / 400)
            Dim DecimalMinutes As Double = ((Gradians * 360 / 400) - Degrees) * 60
            Dim Minutes As Integer = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal = System.Decimal.Round(RawSeconds, 6) 'Round to 6 decimal places
            'Check if RoundedSeconds >= 60
            If RoundedSeconds >= 60 Then 'Rounding has increased the number of seconds past 60.
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes > 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then 'Rounding has decreased the number of (negative) seconds below -60
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes < -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If
            SexagesimalDegrees = Degrees + Minutes / 100 + RoundedSeconds / 10000
        End Sub

        Public Sub ConvertGradianToTurn()
            Turns = Gradians / 400
        End Sub

        Public Sub ConvertGradianToDegMinSec()
            'Convert Gradian to Degrees, Minutes & Seconds
            Dim DecDegrees As Double
            DecDegrees = Gradians * 360 / 400

            Dim AbsDecDegrees As Double
            If DecDegrees < 0 Then
                AbsDecDegrees = Math.Abs(DecDegrees)
                DmsSign = Sign.Negative
            Else
                AbsDecDegrees = DecDegrees
                DmsSign = Sign.Positive
            End If

            DmsDegrees = Fix(AbsDecDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecDegrees - DmsDegrees) * 60
            DmsMinutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - DmsMinutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, DmsSecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                DmsMinutes = DmsMinutes + 1
                RoundedSeconds = 0
                If DmsMinutes >= 60 Then
                    DmsDegrees = DmsDegrees + 1
                    DmsMinutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                DmsMinutes = DmsMinutes - 1
                RoundedSeconds = 0
                If DmsMinutes <= -60 Then
                    DmsDegrees = DmsDegrees - 1
                    DmsMinutes = 0
                End If
            End If
            DmsSeconds = RoundedSeconds
        End Sub

#End Region 'Convert Gradians - Methods to convert gradians to other angle units.

#Region "Convert Deg, Min & Sec"

        Public Sub ConvertDegMinSecToDecimalDegrees()
            'Convert the Degrees, Minutes and Seconds values to a Decimal Degrees value.
            If DmsSign = Sign.Negative Then
                DecimalDegrees = (DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * -1
            Else
                DecimalDegrees = DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600
            End If
        End Sub

        Public Sub ConvertDegMinSecToSexagesimalDegrees()
            'Convert the Degrees, Minutes and Seconds calues to a Sexagesimal Degrees value.
            If DmsSign = Sign.Positive Then
                SexagesimalDegrees = DmsDegrees + DmsMinutes / 100 + DmsSeconds / 10000
            Else
                SexagesimalDegrees = (DmsDegrees + DmsMinutes / 100 + DmsSeconds / 10000) * -1
            End If
        End Sub

        Public Sub ConvertDegMinSecToRadians()
            'Convert the Degrees, Minutes and Seconds values to Radians.
            If DmsSign = Sign.Negative Then
                Radians = ((DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * -1) * 2 * System.Math.PI / 360
            Else
                Radians = (DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * 2 * System.Math.PI / 360
            End If
        End Sub

        Public Sub ConvertDegMinSecToGradians()
            'Convert the Degrees, Minutes and Seconds values to Gradians.
            If DmsSign = Sign.Negative Then
                Gradians = ((DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * -1) * 400 / 360
            Else
                Gradians = (DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * 400 / 360
            End If
        End Sub

        Public Sub ConvertDegMinSecToTurns()
            'Convert the Degrees, Minutes and Seconds values to Turns.
            If DmsSign = Sign.Negative Then
                Turns = ((DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * -1) / 360
            Else
                Turns = (DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) / 360
            End If
        End Sub

#End Region 'Convert Deg, Min & Sec - Methods to convert degrees-minutes-seconds to other angle units.

    End Class 'AngleConvert

    Public Class AngleConvertAuto
        'Converts between decimal degrees, degrees-minutes-seconds, sexagesimal degrees, radians, gradians and turns.
        'Preferred angle units can be selected.
        'Option to automatically convert to preferred angle units.

        Public Enum EnumPreferredAngleUnit
            DecimalDegrees
            DegreesMinutesSeconds
            SexagesimalDegrees
            Radians
            Gradians
            Turns
            NoPreference
        End Enum

        Private _preferredUnit As EnumPreferredAngleUnit = EnumPreferredAngleUnit.NoPreference
        Property PreferredUnit As EnumPreferredAngleUnit 'The preferred angle unit.
            Get
                Return _preferredUnit
            End Get
            Set(value As EnumPreferredAngleUnit)
                _preferredUnit = value
            End Set
        End Property

        Private _autoConvert As Boolean = False
        Property AutoConvert As Boolean 'If True, the input angle will be automatically converted to the preferred unit.
            Get
                Return _autoConvert
            End Get
            Set(value As Boolean)
                _autoConvert = value
            End Set
        End Property

#Region "Angle properties"

        Private _radians As Double = 0
        Property Radians As Double
            Get
                Return _radians
            End Get
            Set(value As Double)
                _radians = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DecimalDegrees
                            ConvertRadianToDecimalDegree()
                        Case EnumPreferredAngleUnit.DegreesMinutesSeconds
                            ConvertRadianToDegMinSec()
                        Case EnumPreferredAngleUnit.Gradians
                            ConvertRadianToGradian()
                        Case EnumPreferredAngleUnit.SexagesimalDegrees
                            ConvertRadianToSexagesimalDegree()
                        Case EnumPreferredAngleUnit.Turns
                            ConvertRadianToTurn()
                    End Select
                End If
            End Set
        End Property

        Private _turns As Double = 0
        Property Turns As Double
            Get
                Return _turns
            End Get
            Set(value As Double)
                _turns = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DecimalDegrees
                            ConvertTurnToDecimalDegree()
                        Case EnumPreferredAngleUnit.DegreesMinutesSeconds
                            ConvertTurnToDegMinSec()
                        Case EnumPreferredAngleUnit.Gradians
                            ConvertTurnToGradian()
                        Case EnumPreferredAngleUnit.Radians
                            ConvertTurnToRadian()
                        Case EnumPreferredAngleUnit.SexagesimalDegrees
                            ConvertTurnToSexagesimalDegree()
                    End Select
                End If
            End Set
        End Property

        Private _decimalDegrees As Double = 0
        Property DecimalDegrees As Double
            Get
                Return _decimalDegrees
            End Get
            Set(value As Double)
                _decimalDegrees = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DegreesMinutesSeconds
                            ConvertDecimalDegreeToDegMinSec()
                        Case EnumPreferredAngleUnit.Gradians
                            ConvertDecimalDegreeToGradian()
                        Case EnumPreferredAngleUnit.Radians
                            ConvertDecimalDegreeToRadian()
                        Case EnumPreferredAngleUnit.SexagesimalDegrees
                            ConvertDecimalDegreeToSexagesimalDegree()
                        Case EnumPreferredAngleUnit.Turns
                            ConvertDecimalDegreeToTurn()
                    End Select
                End If
            End Set
        End Property

        Private _sexagesimalDegrees As Double = 0
        Property SexagesimalDegrees As Double
            Get
                Return _sexagesimalDegrees
            End Get
            Set(value As Double)
                _sexagesimalDegrees = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DecimalDegrees
                            ConvertSexagesimalDegreeToDecimalDegree()
                        Case EnumPreferredAngleUnit.DegreesMinutesSeconds
                            ConvertSexagesimalDegreeToDegMinSec()
                        Case EnumPreferredAngleUnit.Gradians
                            ConvertSexagesimalDegreeToGradian()
                        Case EnumPreferredAngleUnit.Radians
                            ConvertSexagesimalDegreeToRadian()
                        Case EnumPreferredAngleUnit.Turns
                            ConvertSexagesimalDegreeToTurn()
                    End Select
                End If
            End Set
        End Property

        Private _gradians As Double = 0
        Property Gradians As Double
            Get
                Return _gradians
            End Get
            Set(value As Double)
                _gradians = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DecimalDegrees
                            ConvertGradianToDecimalDegree()
                        Case EnumPreferredAngleUnit.DegreesMinutesSeconds
                            ConvertGradianToDegMinSec()
                        Case EnumPreferredAngleUnit.Radians
                            ConvertGradianToRadian()
                        Case EnumPreferredAngleUnit.SexagesimalDegrees
                            ConvertGradianToSexagesimalDegree()
                        Case EnumPreferredAngleUnit.Turns
                            ConvertGradianToTurn()
                    End Select
                End If
            End Set
        End Property

#End Region 'Angle properties: decimal degrees, sexagesimal degrees, radians, gradians & turns.

#Region "Degrees, Minutes & Seconds properties"

        Public Enum Sign
            Positive
            Negative
        End Enum

        Private _dmsSign As Sign = Sign.Positive
        Property DmsSign As Sign
            Get
                Return _dmsSign
            End Get
            Set(value As Sign)
                _dmsSign = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DecimalDegrees
                            ConvertDegMinSecToDecimalDegrees()
                        Case EnumPreferredAngleUnit.Gradians
                            ConvertDegMinSecToGradians()
                        Case EnumPreferredAngleUnit.Radians
                            ConvertDegMinSecToRadians()
                        Case EnumPreferredAngleUnit.SexagesimalDegrees
                            ConvertDegMinSecToSexagesimalDegrees()
                        Case EnumPreferredAngleUnit.Turns
                            ConvertDegMinSecToTurns()
                    End Select
                End If
            End Set
        End Property


        Private _dmsDegrees As Integer = 0
        Property DmsDegrees As Integer
            Get
                Return _dmsDegrees
            End Get
            Set(value As Integer)
                _dmsDegrees = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DecimalDegrees
                            ConvertDegMinSecToDecimalDegrees()
                        Case EnumPreferredAngleUnit.Gradians
                            ConvertDegMinSecToGradians()
                        Case EnumPreferredAngleUnit.Radians
                            ConvertDegMinSecToRadians()
                        Case EnumPreferredAngleUnit.SexagesimalDegrees
                            ConvertDegMinSecToSexagesimalDegrees()
                        Case EnumPreferredAngleUnit.Turns
                            ConvertDegMinSecToTurns()
                    End Select
                End If
            End Set
        End Property

        Private _dmsMinutes As Integer = 0
        Property DmsMinutes As Integer
            Get
                Return _dmsMinutes
            End Get
            Set(value As Integer)
                _dmsMinutes = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DecimalDegrees
                            ConvertDegMinSecToDecimalDegrees()
                        Case EnumPreferredAngleUnit.Gradians
                            ConvertDegMinSecToGradians()
                        Case EnumPreferredAngleUnit.Radians
                            ConvertDegMinSecToRadians()
                        Case EnumPreferredAngleUnit.SexagesimalDegrees
                            ConvertDegMinSecToSexagesimalDegrees()
                        Case EnumPreferredAngleUnit.Turns
                            ConvertDegMinSecToTurns()
                    End Select
                End If
            End Set
        End Property


        Private _dmsSeconds As Decimal = 0
        Property DmsSeconds As Decimal
            Get
                Return _dmsSeconds
            End Get
            Set(value As Decimal)
                _dmsSeconds = value
                If _autoConvert = True Then
                    Select Case _preferredUnit
                        Case EnumPreferredAngleUnit.NoPreference
                            'No preferred unit to convert to.
                        Case EnumPreferredAngleUnit.DecimalDegrees
                            ConvertDegMinSecToDecimalDegrees()
                        Case EnumPreferredAngleUnit.Gradians
                            ConvertDegMinSecToGradians()
                        Case EnumPreferredAngleUnit.Radians
                            ConvertDegMinSecToRadians()
                        Case EnumPreferredAngleUnit.SexagesimalDegrees
                            ConvertDegMinSecToSexagesimalDegrees()
                        Case EnumPreferredAngleUnit.Turns
                            ConvertDegMinSecToTurns()
                    End Select
                End If
            End Set
        End Property

        Private _dmsSecondsDecimalPlaces As Integer = 4
        Property DmsSecondsDecimalPlaces As Integer 'The number of decimal places used to display the seconds value.
            Get
                Return _dmsSecondsDecimalPlaces
            End Get
            Set(value As Integer)
                _dmsSecondsDecimalPlaces = value
            End Set
        End Property

#End Region 'Degrees, Minutes & Seconds properties

#Region "Convert Radians"
        'Radians to: Decimal degrees, Sexagesimal degrees, Turns, Gradians -----------------------------------------------------

        Public Sub ConvertRadianToDecimalDegree()
            DecimalDegrees = Radians * 360 / 2 / System.Math.PI
        End Sub

        Public Sub ConvertRadianToSexagesimalDegree()
            Dim Degrees As Integer = Fix(Radians * 360 / 2 / System.Math.PI)
            Dim DecimalMinutes As Double = (Radians * 360 / 2 / System.Math.PI - Degrees) * 60
            Dim Minutes As Integer = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal = System.Decimal.Round(RawSeconds, 6) 'Round to 6 decimal places
            'Check if RoundedSeconds >= 60
            If RoundedSeconds >= 60 Then 'Rounding has increased the number of seconds past 60.
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes > 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then 'Rounding has decreased the number of (negative) seconds below -60
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes < -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If
            SexagesimalDegrees = Degrees + Minutes / 100 + RoundedSeconds / 10000
        End Sub

        Public Sub ConvertRadianToTurn()
            Turns = Radians / 2 / System.Math.PI
        End Sub

        Public Sub ConvertRadianToGradian()
            Gradians = Radians * 400 / 2 / System.Math.PI
        End Sub

        Public Sub ConvertRadianToDegMinSec()
            'Convert Radian to Degrees, Minutes & Seconds
            Dim DecDegrees As Double
            DecDegrees = Radians * 360 / 2 / System.Math.PI

            Dim AbsDecDegrees As Double
            If DecDegrees < 0 Then
                AbsDecDegrees = Math.Abs(DecDegrees)
                DmsSign = Sign.Negative
            Else
                AbsDecDegrees = DecDegrees
                DmsSign = Sign.Positive
            End If

            DmsDegrees = Fix(AbsDecDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecDegrees - DmsDegrees) * 60
            DmsMinutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - DmsMinutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, DmsSecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                DmsMinutes = DmsMinutes + 1
                RoundedSeconds = 0
                If DmsMinutes >= 60 Then
                    DmsDegrees = DmsDegrees + 1
                    DmsMinutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                DmsMinutes = DmsMinutes - 1
                RoundedSeconds = 0
                If DmsMinutes <= -60 Then
                    DmsDegrees = DmsDegrees - 1
                    DmsMinutes = 0
                End If
            End If
            DmsSeconds = RoundedSeconds
        End Sub

#End Region 'Convert Radians - Methods to convert radians to other angle units.

#Region "Convert Decimal Degrees"
        'Decimal degrees to: Radians, Sexagesimal degrees, Turns, Gradians -------------------------------------------------------

        Public Sub ConvertDecimalDegreeToRadian()
            Radians = DecimalDegrees * 2 * System.Math.PI / 360
        End Sub

        Public Sub ConvertDecimalDegreeToSexagesimalDegree()
            Dim Degrees As Integer = Fix(DecimalDegrees)
            Dim DecimalMinutes As Double = (DecimalDegrees - Degrees) * 60
            Dim Minutes As Integer = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal = System.Decimal.Round(RawSeconds, 6) 'Round to 6 decimal places
            'Check if RoundedSeconds >= 60
            If RoundedSeconds >= 60 Then 'Rounding has increased the number of seconds past 60.
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes > 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then 'Rounding has decreased the number of (negative) seconds below -60
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes < -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If
            SexagesimalDegrees = Degrees + Minutes / 100 + RoundedSeconds / 10000
        End Sub

        Public Sub ConvertDecimalDegreeToTurn()
            Turns = DecimalDegrees / 360
        End Sub

        Public Sub ConvertDecimalDegreeToGradian()
            Gradians = DecimalDegrees * 400 / 360
        End Sub

        Public Sub ConvertDecimalDegreeToDegMinSec()
            'Convert DecimalDegree to Degrees, Minutes & Seconds
            Dim AbsDecimalDegrees As Double
            If DecimalDegrees < 0 Then
                AbsDecimalDegrees = Math.Abs(DecimalDegrees)
                DmsSign = Sign.Negative
            Else
                AbsDecimalDegrees = DecimalDegrees
                DmsSign = Sign.Positive
            End If

            DmsDegrees = Fix(AbsDecimalDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecimalDegrees - DmsDegrees) * 60
            DmsMinutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - DmsMinutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, DmsSecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                DmsMinutes = DmsMinutes + 1
                RoundedSeconds = 0
                If DmsMinutes >= 60 Then
                    DmsDegrees = DmsDegrees + 1
                    DmsMinutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                DmsMinutes = DmsMinutes - 1
                RoundedSeconds = 0
                If DmsMinutes <= -60 Then
                    DmsDegrees = DmsDegrees - 1
                    DmsMinutes = 0
                End If
            End If
            DmsSeconds = RoundedSeconds
        End Sub

#End Region 'Convert Decimal Degrees - Methods to convert decimal degrees to other angle units.


#Region "Convert Sexagesimal Degrees"
        'Sexagesimal degrees to: Radians, Decimal degrees, Turns, Gradians ------------------------------------------------------

        Public Sub ConvertSexagesimalDegreeToRadian()
            'Dim SexagesimalStr As String = Trim(Str(SexagesimalDegrees))
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            Dim Minutes As Integer = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            Dim Seconds As Double = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
            If SexagesimalStr.StartsWith("-") Then
                Radians = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * -2 * System.Math.PI / 360
            ElseIf SexagesimalStr.StartsWith("+") Then
                Radians = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * 2 * System.Math.PI / 360
            Else
                Radians = (SexagesimalStr.Substring(0, DecimalPointPos) + Minutes / 60 + Seconds / 3600) * 2 * System.Math.PI / 360
            End If
        End Sub


        Public Sub ConvertSexagesimalDegreeToDecimalDegree()
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            Dim Minutes As Integer = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            Dim Seconds As Double = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
            If SexagesimalStr.StartsWith("-") Then
                DecimalDegrees = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * -1
            ElseIf SexagesimalStr.StartsWith("+") Then
                DecimalDegrees = SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600
            Else
                DecimalDegrees = SexagesimalStr.Substring(0, DecimalPointPos) + Minutes / 60 + Seconds / 3600
            End If
        End Sub


        Public Sub ConvertSexagesimalDegreeToTurn()
            'Dim SexagesimalStr As String = Trim(Str(SexagesimalDegrees))
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            Dim Minutes As Integer = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            Dim Seconds As Double = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
            If SexagesimalStr.StartsWith("-") Then
                Turns = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * -1 / 360
            ElseIf SexagesimalStr.StartsWith("+") Then
                Turns = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) / 360
            Else
                Turns = (SexagesimalStr.Substring(0, DecimalPointPos) + Minutes / 60 + Seconds / 3600) / 360
            End If
        End Sub


        Public Sub ConvertSexagesimalDegreeToGradian()
            'Dim SexagesimalStr As String = Trim(Str(SexagesimalDegrees))
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            Dim Minutes As Integer = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            Dim Seconds As Double = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
            If SexagesimalStr.StartsWith("-") Then
                Gradians = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * -400 / 360
            ElseIf SexagesimalStr.StartsWith("+") Then
                Gradians = (SexagesimalStr.Substring(1, DecimalPointPos - 1) + Minutes / 60 + Seconds / 3600) * 400 / 360
            Else
                Gradians = (SexagesimalStr.Substring(0, DecimalPointPos) + Minutes / 60 + Seconds / 3600) * 400 / 360
            End If
        End Sub


        Public Sub ConvertSexagesimalDegreeToDegMinSec()
            'To avoid errors caused by rounding, Sexagesimal degrees are converted using character positions in the Sexagesimal string.
            Dim SexagesimalStr As String = Format(SexagesimalDegrees, "##0.0000##############")
            Dim DecimalPointPos As Integer = SexagesimalStr.IndexOf(".") 'The (zero based) character position of the decimal point
            If SexagesimalStr.StartsWith("-") Then
                DmsSign = Sign.Negative
                DmsDegrees = SexagesimalStr.Substring(1, DecimalPointPos - 1)
            ElseIf SexagesimalStr.StartsWith("+") Then
                DmsSign = Sign.Positive
                DmsDegrees = SexagesimalStr.Substring(1, DecimalPointPos - 1)
            Else
                DmsSign = Sign.Positive
                DmsDegrees = SexagesimalStr.Substring(0, DecimalPointPos)
            End If
            DmsMinutes = SexagesimalStr.Substring(DecimalPointPos + 1, 2) 'Selects the two characters past the decimal point
            DmsSeconds = SexagesimalStr.Substring(DecimalPointPos + 3, SexagesimalStr.Length - DecimalPointPos - 3)
        End Sub

#End Region 'Convert Sexagesimal Degrees - Methods to convert sexagesimal degrees to other angle units.


#Region "Convert Turns"
        'Turns to: Radian, Decimal degrees, Sexagesimal degrees, Gradians -------------------------------------------------------

        Public Sub ConvertTurnToRadian()
            Radians = Turns * 2 * System.Math.PI
        End Sub

        Public Sub ConvertTurnToDecimalDegree()
            DecimalDegrees = Turns * 360
        End Sub

        Public Sub ConvertTurnToSexagesimalDegree()
            Dim Degrees As Integer = Fix(Turns * 360)
            Dim DecimalMinutes As Double = (Turns * 360 - Degrees) * 60
            Dim Minutes As Integer = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal = System.Decimal.Round(RawSeconds, 6) 'Round to 6 decimal places
            'Check if RoundedSeconds >= 60
            If RoundedSeconds >= 60 Then 'Rounding has increased the number of seconds past 60.
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes > 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then 'Rounding has decreased the number of (negative) seconds below -60
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes < -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If
            SexagesimalDegrees = Degrees + Minutes / 100 + RoundedSeconds / 10000
        End Sub

        Public Sub ConvertTurnToGradian()
            Gradians = Turns * 400
        End Sub

        Public Sub ConvertTurnToDegMinSec()
            'Convert Turn to Degrees, Minutes & Seconds
            Dim DecDegrees As Double
            DecDegrees = Turns * 360

            Dim AbsDecDegrees As Double
            If DecDegrees < 0 Then
                AbsDecDegrees = Math.Abs(DecDegrees)
                DmsSign = Sign.Negative
            Else
                AbsDecDegrees = DecDegrees
                DmsSign = Sign.Positive
            End If

            DmsDegrees = Fix(AbsDecDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecDegrees - DmsDegrees) * 60
            DmsMinutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - DmsMinutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, DmsSecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                DmsMinutes = DmsMinutes + 1
                RoundedSeconds = 0
                If DmsMinutes >= 60 Then
                    DmsDegrees = DmsDegrees + 1
                    DmsMinutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                DmsMinutes = DmsMinutes - 1
                RoundedSeconds = 0
                If DmsMinutes <= -60 Then
                    DmsDegrees = DmsDegrees - 1
                    DmsMinutes = 0
                End If
            End If
            DmsSeconds = RoundedSeconds
        End Sub

#End Region 'Convert Turns - Methods to convert turns to other angle units.


#Region "Convert Gradians"

        'Gradians to: Radians, Decimal degrees, Sexagesimal degrees, Turns ------------------------------------------------------
        Public Sub ConvertGradianToRadian()
            Radians = Gradians * 2 * System.Math.PI / 400
        End Sub

        Public Sub ConvertGradianToDecimalDegree()
            DecimalDegrees = Gradians * 360 / 400
        End Sub

        Public Sub ConvertGradianToSexagesimalDegree()
            Dim Degrees As Integer = Fix(Gradians * 360 / 400)
            Dim DecimalMinutes As Double = ((Gradians * 360 / 400) - Degrees) * 60
            Dim Minutes As Integer = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal = (DecimalMinutes - Minutes) * 60
            Dim RoundedSeconds As Decimal = System.Decimal.Round(RawSeconds, 6) 'Round to 6 decimal places
            'Check if RoundedSeconds >= 60
            If RoundedSeconds >= 60 Then 'Rounding has increased the number of seconds past 60.
                Minutes = Minutes + 1
                RoundedSeconds = 0
                If Minutes > 60 Then
                    Degrees = Degrees + 1
                    Minutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then 'Rounding has decreased the number of (negative) seconds below -60
                Minutes = Minutes - 1
                RoundedSeconds = 0
                If Minutes < -60 Then
                    Degrees = Degrees - 1
                    Minutes = 0
                End If
            End If
            SexagesimalDegrees = Degrees + Minutes / 100 + RoundedSeconds / 10000
        End Sub

        Public Sub ConvertGradianToTurn()
            Turns = Gradians / 400
        End Sub

        Public Sub ConvertGradianToDegMinSec()
            'Convert Gradian to Degrees, Minutes & Seconds
            Dim DecDegrees As Double
            DecDegrees = Gradians * 360 / 400

            Dim AbsDecDegrees As Double
            If DecDegrees < 0 Then
                AbsDecDegrees = Math.Abs(DecDegrees)
                DmsSign = Sign.Negative
            Else
                AbsDecDegrees = DecDegrees
                DmsSign = Sign.Positive
            End If

            DmsDegrees = Fix(AbsDecDegrees)
            Dim DecimalMinutes As Double
            DecimalMinutes = (AbsDecDegrees - DmsDegrees) * 60
            DmsMinutes = Fix(DecimalMinutes)
            Dim RawSeconds As Decimal
            RawSeconds = (DecimalMinutes - DmsMinutes) * 60
            Dim RoundedSeconds As Decimal
            RoundedSeconds = System.Decimal.Round(RawSeconds, DmsSecondsDecimalPlaces)
            'Check if RoundedSeconds > 60
            If RoundedSeconds >= 60 Then
                DmsMinutes = DmsMinutes + 1
                RoundedSeconds = 0
                If DmsMinutes >= 60 Then
                    DmsDegrees = DmsDegrees + 1
                    DmsMinutes = 0
                End If
            ElseIf RoundedSeconds <= -60 Then
                DmsMinutes = DmsMinutes - 1
                RoundedSeconds = 0
                If DmsMinutes <= -60 Then
                    DmsDegrees = DmsDegrees - 1
                    DmsMinutes = 0
                End If
            End If
            DmsSeconds = RoundedSeconds
        End Sub

#End Region 'Convert Gradians - Methods to convert gradians to other angle units.


#Region "Convert Deg, Min & Sec"

        Public Sub ConvertDegMinSecToDecimalDegrees()
            'Convert the Degrees, Minutes and Seconds values to a Decimal Degrees value.
            If DmsSign = Sign.Negative Then
                DecimalDegrees = (DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * -1
            Else
                DecimalDegrees = DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600
            End If
        End Sub

        Public Sub ConvertDegMinSecToSexagesimalDegrees()
            'Convert the Degrees, Minutes and Seconds calues to a Sexagesimal Degrees value.
            If DmsSign = Sign.Positive Then
                SexagesimalDegrees = DmsDegrees + DmsMinutes / 100 + DmsSeconds / 10000
            Else
                SexagesimalDegrees = (DmsDegrees + DmsMinutes / 100 + DmsSeconds / 10000) * -1
            End If
        End Sub

        Public Sub ConvertDegMinSecToRadians()
            'Convert the Degrees, Minutes and Seconds values to Radians.
            If DmsSign = Sign.Negative Then
                Radians = ((DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * -1) * 2 * System.Math.PI / 360
            Else
                Radians = (DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * 2 * System.Math.PI / 360
            End If
        End Sub

        Public Sub ConvertDegMinSecToGradians()
            'Convert the Degrees, Minutes and Seconds values to Gradians.
            If DmsSign = Sign.Negative Then
                Gradians = ((DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * -1) * 400 / 360
            Else
                Gradians = (DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * 400 / 360
            End If
        End Sub

        Public Sub ConvertDegMinSecToTurns()
            'Convert the Degrees, Minutes and Seconds values to Turns.
            If DmsSign = Sign.Negative Then
                Turns = ((DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) * -1) / 360
            Else
                Turns = (DmsDegrees + DmsMinutes / 60 + DmsSeconds / 3600) / 360
            End If
        End Sub

#End Region 'Convert Deg, Min & Sec - Methods to convert degrees-minutes-seconds to other angle units.


    End Class 'AngleConvertAuto

    Public Class Location
        'The location class is used to store a point location in different forms.

        'The geodetic latitude coordinate
        Private _latitude As Double
        Property Latitude As Double
            Get
                Return _latitude
            End Get
            Set(value As Double)
                _latitude = value
            End Set
        End Property

        'The geodetic longitude coordinate
        Private _longitude As Double
        Property Longitude As Double
            Get
                Return _longitude
            End Get
            Set(value As Double)
                _longitude = value
            End Set
        End Property

        'The height relative to the datum ellipsoid
        Private _ellipsoidalHeight
        Property EllipsoidalHeight As Double
            Get
                Return _ellipsoidalHeight
            End Get
            Set(value As Double)
                _ellipsoidalHeight = value
            End Set
        End Property

        'The cartesian X coordinate
        Private _x As Double
        Property X As Double
            Get
                Return _x
            End Get
            Set(value As Double)
                _x = value
            End Set
        End Property

        'The cartesian Y coordinate
        Private _y As Double
        Property Y As Double
            Get
                Return _y
            End Get
            Set(value As Double)
                _y = value
            End Set
        End Property

        'The cartesian Z coordinate
        Private _z As Double
        Property Z As Double
            Get
                Return _z
            End Get
            Set(value As Double)
                _z = value
            End Set
        End Property

        'The projected easting coordinate
        Private _easting As Double
        Property Easting As Double
            Get
                Return _easting
            End Get
            Set(value As Double)
                _easting = value
            End Set
        End Property

        'The projected northing coordinate
        Private _northing As Double
        Property Northing As Double
            Get
                Return _northing
            End Get
            Set(value As Double)
                _northing = value
            End Set
        End Property

    End Class 'Location



    'TO DO integrate ListInfo class with the ADVL_System_Utilties project management classes. ===============================================================================================
    'DEFINE THIS CLASS IN THE ADVL_Coordinates application!
    'THE Project class in the ADVL_System_Utilties library can then be used to handle project storage!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'NOTE: ListInfo WILL BE REPLACED BY THE PARAMETER LIST CLASSES. THESE CLASSES WILL INCLUDE METHODS TO READ AND WRITE DATA BETWEEN THE LIST AND XML FILES IN DISK.

    'Public Class ListInfo
    '    'This class is used to hold information about lists of datum information.
    '    'Different datums, projections and other coordinate information are stored in lists of parameters in XML files.
    '    'Applications may use a Project Directory or a Project File to store data.
    '    'Information about any Project Directories or Project Files are included in this class.

    '    'ListInfo PROPERTIES:
    '    'PathType                 (Directory, ProjectFile)     changed from: (ProjectDirectory, GeneralDirectory, ProjectFile, GeneralProjectFile)
    '    'DirectoryPath
    '    'IsProjectDir             (True, False)
    '    'ProjectFileName
    '    'ListFileName
    '    'CreationDate
    '    'LastEditDate
    '    'Description
    '    'NRecords
    '    'NUsers      

    '    'Properties removed:
    '    'ProjectFileInProjectDir (True, False)
    '    'ProjectFileDirPath

    '    'TO DO: Update List Info properties!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    '    'Public Enum enumPathType
    '    '    ProjectDirectory   'The Project Directory is used to store the list file.
    '    '    GeneralDirectory   'A general directory is used to store the list file.
    '    '    ProjectFile        'A Project File in the Project Directory is used to store the list file.
    '    '    GeneralProjectFile 'A Project File in a specified directory is used to store the list file.
    '    'End Enum

    '    Public Enum enumPathType
    '        Directory          'The list file is stored in a Directory (Project Directory or other directory).
    '        ProjectFile        'The list file is stored in a Project File.
    '    End Enum

    '    Private _pathType As enumPathType = enumPathType.Directory
    '    Property PathType As enumPathType
    '        Get
    '            Return _pathType
    '        End Get
    '        Set(value As enumPathType)
    '            _pathType = value
    '        End Set
    '    End Property

    '    'TO BE REMOVED: (FilePath can be constructed from the PathType, DirectoryPath, ListFileName and ProjectFileName properties)
    '    Private _FilePath As String = ""
    '    Property FilePath As String 'The path to the list file.
    '        'If the PathType is ProjectDirectory or GeneralDirectory then the FilePath is used to access the list file.
    '        Get
    '            Return _FilePath
    '        End Get
    '        Set(value As String)
    '            _FilePath = value
    '        End Set
    '    End Property

    '    Private _projectFileName As String = ""
    '    Property ProjectFileName As String
    '        Get
    '            Return _projectFileName
    '        End Get
    '        Set(value As String)
    '            _projectFileName = value
    '        End Set
    '    End Property

    '    Private _listFileName As String = ""
    '    Property ListFileName As String 'The file name (with extension) of the list file.
    '        'If the PathType is ProjectFile then the ProjectFilePath and FileName is used to access the list file.
    '        Get
    '            Return _listFileName
    '        End Get
    '        Set(value As String)
    '            _listFileName = value
    '        End Set
    '    End Property


    '    Private _directoryPath As String = ""
    '    Property DirectoryPath As String 'The path the the Project File if it is used to store the list file.
    '        'If the PathType is ProjectFile or GeneralProjectFile then the ProjectFilePath and FileName is used to access the list file.
    '        Get
    '            Return _directoryPath
    '        End Get
    '        Set(value As String)
    '            _directoryPath = value
    '        End Set
    '    End Property


    '    Private _isProjectDir As Boolean = True
    '    Property IsProjectDir As Boolean 'True if the DirectoryPath corresponds to the Project Directory
    '        Get
    '            Return _isProjectDir
    '        End Get
    '        Set(value As Boolean)
    '            _isProjectDir = value
    '        End Set
    '    End Property

    '    Private _creationDate As DateTime = Now
    '    Property CreationDate As DateTime
    '        Get
    '            Return _creationDate
    '        End Get
    '        Set(value As DateTime)
    '            _creationDate = value
    '        End Set
    '    End Property

    '    Private _lastEditDate As DateTime = Now
    '    Property LastEditDate As DateTime
    '        Get
    '            Return _lastEditDate
    '        End Get
    '        Set(value As DateTime)
    '            _lastEditDate = value
    '        End Set
    '    End Property


    '    Private _description As String = ""
    '    Property Description As String
    '        Get
    '            Return _description
    '        End Get
    '        Set(value As String)
    '            _description = value
    '        End Set
    '    End Property

    '    Private _nRecords As Integer = 0
    '    Property NRecords As Integer
    '        Get
    '            Return _nRecords
    '        End Get
    '        Set(value As Integer)
    '            _nRecords = value
    '            'If _currentRecordNo > _nRecords Then
    '            '    CurrentRecordNo = NRecords
    '            'End If
    '        End Set
    '    End Property

    '    Private _nUsers As Integer = 0
    '    Property NUsers As Integer 'The number of users connected ot the list.
    '        Get
    '            Return _nUsers
    '        End Get
    '        Set(value As Integer)
    '            _nUsers = value
    '        End Set
    '    End Property


    '    Public Sub AddUser()
    '        'Add a user to the list
    '        _nUsers += 1
    '    End Sub

    '    Public Sub RemoveUser()
    '        'Remove a user for the list.
    '        If _nUsers > 0 Then
    '            _nUsers -= 1
    '        End If
    '    End Sub


    'End Class 'ListInfo

    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!



#Region " Lists of Parameters" '--------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class AreaOfUseList
        'Most coordinate systems have a specified area of use. 
        'This class contains a list of these areas.

        Public List As New List(Of AreaOfUse) 'List of Area Of Use parameters.

    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation 'The location of the list file.

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property

        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

        Private _findIndexResult As FindIndexResults = FindIndexResults.NotFound
        Property FindIndexResult As FindIndexResults 'The result of searching for a specified record in the list.
            Get
                Return _findIndexResult
            End Get
            Set(value As FindIndexResults)
                _findIndexResult = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the AreaOfUse List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<AOUList>.<CreationDate>.Value
            LastEditDate = XDoc.<AOUList>.<LastEditDate>.Value
            Description = XDoc.<AOUList>.<Description>.Value

            Dim AOUs = From item In XDoc.<AOUList>.<AOU>

            List.Clear()
            For Each item In AOUs
                Dim NewAOU As New AreaOfUse
                NewAOU.Name = item.<Name>.Value
                NewAOU.Author = item.<Author>.Value
                If NewAOU.Author = Nothing Then NewAOU.Author = ""
                NewAOU.Code = item.<Code>.Value
                If NewAOU.Code = Nothing Then NewAOU.Code = 0
                NewAOU.Comments = item.<Comments>.Value

                If item.<Deprecated>.Value = Nothing Then
                    NewAOU.Deprecated = True
                Else
                    NewAOU.Deprecated = item.<Deprecated>.Value
                End If

                NewAOU.Description = item.<Description>.Value
                If item.<SouthLatitude>.Value = "" Then
                    NewAOU.SouthLatitude = 0
                Else
                    NewAOU.SouthLatitude = item.<SouthLatitude>.Value
                End If
                If item.<NorthLatitude>.Value = "" Then
                    NewAOU.NorthLatitude = 0
                Else
                    NewAOU.NorthLatitude = item.<NorthLatitude>.Value
                End If
                If item.<WestLongitude>.Value = "" Then
                    NewAOU.WestLongitude = 0
                Else
                    NewAOU.WestLongitude = item.<WestLongitude>.Value
                End If
                If item.<EastLongitude>.Value = "" Then
                    NewAOU.EastLongitude = 0
                Else
                    NewAOU.EastLongitude = item.<EastLongitude>.Value
                End If

                NewAOU.IsoA2Code = item.<IsoA2Code>.Value
                NewAOU.IsoA3Code = item.<IsoA3Code>.Value
                If item.<IsoNCode>.Value = "" Then
                    NewAOU.IsoNCode = 0
                Else
                    NewAOU.IsoNCode = item.<IsoNCode>.Value
                End If
                Dim aliasNames = From item2 In item.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewAOU.AddAlias(item3)
                Next
                List.Add(NewAOU)
            Next

        End Sub

        'Load the list from the selected list file.
        Public Sub LoadFile()
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If IsNothing(XDoc) Then
                Exit Sub
            End If

            CreationDate = XDoc.<AOUList>.<CreationDate>.Value
            LastEditDate = XDoc.<AOUList>.<LastEditDate>.Value
            Description = XDoc.<AOUList>.<Description>.Value

            Dim AOUs = From item In XDoc.<AOUList>.<AOU>

            List.Clear()
            For Each item In AOUs
                Dim NewAOU As New AreaOfUse
                NewAOU.Name = item.<Name>.Value
                NewAOU.Author = item.<Author>.Value
                If NewAOU.Author = Nothing Then NewAOU.Author = ""
                NewAOU.Code = item.<Code>.Value
                If NewAOU.Code = Nothing Then NewAOU.Code = 0
                NewAOU.Comments = item.<Comments>.Value

                If item.<Deprecated>.Value = Nothing Then
                    NewAOU.Deprecated = True
                Else
                    NewAOU.Deprecated = item.<Deprecated>.Value
                End If

                NewAOU.Description = item.<Description>.Value
                If item.<SouthLatitude>.Value = "" Then
                    NewAOU.SouthLatitude = 0
                Else
                    NewAOU.SouthLatitude = item.<SouthLatitude>.Value
                End If
                If item.<NorthLatitude>.Value = "" Then
                    NewAOU.NorthLatitude = 0
                Else
                    NewAOU.NorthLatitude = item.<NorthLatitude>.Value
                End If
                If item.<WestLongitude>.Value = "" Then
                    NewAOU.WestLongitude = 0
                Else
                    NewAOU.WestLongitude = item.<WestLongitude>.Value
                End If
                If item.<EastLongitude>.Value = "" Then
                    NewAOU.EastLongitude = 0
                Else
                    NewAOU.EastLongitude = item.<EastLongitude>.Value
                End If

                NewAOU.IsoA2Code = item.<IsoA2Code>.Value
                NewAOU.IsoA3Code = item.<IsoA3Code>.Value
                If item.<IsoNCode>.Value = "" Then
                    NewAOU.IsoNCode = 0
                Else
                    NewAOU.IsoNCode = item.<IsoNCode>.Value
                End If
                Dim aliasNames = From item2 In item.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewAOU.AddAlias(item3)
                Next
                List.Add(NewAOU)
            Next

        End Sub

        'Load the Area Of Use list from the EPSG database.
        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            ds.Clear()
            ds.Reset()

            da = New OleDb.OleDbDataAdapter("Select * From [Area]", myConnection)
            TableName = "[Area]"
            da.Fill(ds, TableName) 'Read the Area table into ds

            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "[Alias]"
            da.Fill(ds, TableName) 'Read the Alias table into ds

            'Dim AOUCode As Integer
            Dim expression As String

            Dim NCols As Integer = ds.Tables(0).Columns.Count
            Dim NRows As Integer = ds.Tables(0).Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Area Of Use list"

            Dim RowNo As Integer

            For RowNo = 0 To NRows - 1 'loop through each row in the AOUs table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewAou As New AreaOfUse
                NewAou.Name = ds.Tables("[Area]").Rows(RowNo).Item("AREA_NAME")
                NewAou.Author = "EPSG"
                NewAou.Code = ds.Tables("[Area]").Rows(RowNo).Item("AREA_CODE")
                If IsDBNull(ds.Tables("[Area]").Rows(RowNo).Item("REMARKS")) Then
                    NewAou.Comments = ""
                Else
                    NewAou.Comments = ds.Tables("[Area]").Rows(RowNo).Item("REMARKS")
                End If

                NewAou.Deprecated = ds.Tables("[Area]").Rows(RowNo).Item("DEPRECATED")
                NewAou.Description = ds.Tables("[Area]").Rows(RowNo).Item("AREA_OF_USE")
                If IsDBNull(ds.Tables("[Area]").Rows(RowNo).Item("AREA_SOUTH_BOUND_LAT")) Then
                    NewAou.SouthLatitude = 0
                Else
                    NewAou.SouthLatitude = ds.Tables("[Area]").Rows(RowNo).Item("AREA_SOUTH_BOUND_LAT")
                End If
                If IsDBNull(ds.Tables("[Area]").Rows(RowNo).Item("AREA_NORTH_BOUND_LAT")) Then
                    NewAou.NorthLatitude = 0
                Else
                    NewAou.NorthLatitude = ds.Tables("[Area]").Rows(RowNo).Item("AREA_NORTH_BOUND_LAT")
                End If
                If IsDBNull(ds.Tables("[Area]").Rows(RowNo).Item("AREA_WEST_BOUND_LON")) Then
                    NewAou.WestLongitude = 0
                Else
                    NewAou.WestLongitude = ds.Tables("[Area]").Rows(RowNo).Item("AREA_WEST_BOUND_LON")
                End If
                If IsDBNull(ds.Tables("[Area]").Rows(RowNo).Item("AREA_EAST_BOUND_LON")) Then
                    NewAou.EastLongitude = 0
                Else
                    NewAou.EastLongitude = ds.Tables("[Area]").Rows(RowNo).Item("AREA_EAST_BOUND_LON")
                End If

                If IsDBNull(ds.Tables("[Area]").Rows(RowNo).Item("ISO_A2_CODE")) Then
                    NewAou.IsoA2Code = ""
                Else
                    NewAou.IsoA2Code = ds.Tables("[Area]").Rows(RowNo).Item("ISO_A2_CODE")
                End If
                If IsDBNull(ds.Tables("[Area]").Rows(RowNo).Item("ISO_A3_CODE")) Then
                    NewAou.IsoA3Code = ""
                Else
                    NewAou.IsoA3Code = ds.Tables("[Area]").Rows(RowNo).Item("ISO_A3_CODE")
                End If
                If IsDBNull(ds.Tables("[Area]").Rows(RowNo).Item("ISO_N_CODE")) Then
                    NewAou.IsoNCode = 0
                Else
                    NewAou.IsoNCode = ds.Tables("[Area]").Rows(RowNo).Item("ISO_N_CODE")
                End If

                expression = "[OBJECT_TABLE_NAME] = 'Area' AND [OBJECT_CODE] = " & Str(NewAou.Code)

                Dim result = ds.Tables("[Alias]").Select(expression)

                For Each item In result
                    NewAou.AddAlias(item.Item("ALIAS").ToString)
                Next

                List.Add(NewAou)
            Next

            myConnection.Close()
        End Sub

        'Function to return the list of Areas of Use as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Areas Of Use List File-->
                       <AOUList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <AOU>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Comments><%= item.Comments %></Comments>
                                     <Description><%= item.Description %></Description>
                                     <SouthLatitude><%= item.SouthLatitude %></SouthLatitude>
                                     <NorthLatitude><%= item.NorthLatitude %></NorthLatitude>
                                     <WestLongitude><%= item.WestLongitude %></WestLongitude>
                                     <EastLongitude><%= item.EastLongitude %></EastLongitude>
                                     <IsoA2Code><%= item.IsoA2Code %></IsoA2Code>
                                     <IsoA3Code><%= item.IsoA3Code %></IsoA3Code>
                                     <IsoNCode><%= item.IsoNCode %></IsoNCode>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                 </AOU>
                           %>
                       </AOUList>

            Return XDoc

        End Function

        'Extract an Area Of Use list from the EPSG Access database.
        Public Sub GetEpsgList(ByVal EpsgDatabasePath As String, ByRef XDoc As System.Xml.Linq.XDocument)
            'Extract an Area Of Use list from the EPSG Access database.
            'The database path is EpsgDatabasePath.
            'The list is converted to an XDocument in XDoc.

            'NOTE: LoadEpsgDbList imports AOUs directly into the list.

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = "[Area]"

            If EpsgDatabasePath = "" Then
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) Then

            Else 'Database not found!
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            da = New OleDb.OleDbDataAdapter("Select * From [Area]", myConnection)

            ds.Clear()
            ds.Reset()

            da.Fill(ds, TableName) 'Read the Area table into ds

            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "[Alias]"
            da.Fill(ds, TableName) 'Read the Alias table into ds

            Dim AOUCode As Integer
            Dim expression As String

            Dim NCols As Integer = ds.Tables(0).Columns.Count
            Dim NRows As Integer = ds.Tables(0).Rows.Count

            XDoc.Add(New XComment(""))
            XDoc.Add(New XComment("Areas Of Use List File"))

            Dim aouList As New XElement("AOUList")
            Dim creationDate As New XElement("CreationDate", Format(Now, "d-MMM-yyyy H:mm:ss"))
            aouList.Add(creationDate)
            Dim lastEditDate As New XElement("LastEditDate", Format(Now, "d-MMM-yyyy H:mm:ss"))
            aouList.Add(lastEditDate)
            Dim description As New XElement("Description", "Area Of Use List File")
            aouList.Add(description)

            Dim aous As New XElement("AOUs")

            Dim RowNo As Integer
            Dim Count As Integer = 0
            For RowNo = 0 To NRows - 1 'loop through each row in the AOUs table:
                Dim aou As New XElement("AOU")
                Dim name As New XElement("Name", ds.Tables(0).Rows(RowNo).Item("AREA_NAME"))
                aou.Add(name)
                Count += 1
                If Count Mod 100 = 0 Then
                    'Main.MessageAdd("Reading item " & Count & " of " & NRows & vbCrLf)
                    RaiseEvent Message("Reading item " & Count & " of " & NRows & vbCrLf)
                End If
                Dim author As New XElement("Author", "EPSG")
                aou.Add(author)
                Dim Code As New XElement("Code", ds.Tables(0).Rows(RowNo).Item("AREA_CODE"))
                aou.Add(Code)
                Dim comments As New XElement("Comments", ds.Tables(0).Rows(RowNo).Item("REMARKS"))
                aou.Add(comments)
                'Dim areaOfUse As New XElement("AreaOfUse", ds.Tables(0).Rows(RowNo).Item("AREA_OF_USE"))
                'aou.Add(areaOfUse)
                Dim deprecated As New XElement("Deprecated", ds.Tables(0).Rows(RowNo).Item("DEPRECATED"))
                aou.Add(deprecated)
                Dim aouDescription As New XElement("Description", ds.Tables(0).Rows(RowNo).Item("AREA_OF_USE"))
                aou.Add(aouDescription)
                Dim southLatitude As New XElement("SouthLatitude", ds.Tables(0).Rows(RowNo).Item("AREA_SOUTH_BOUND_LAT"))
                aou.Add(southLatitude)
                Dim northLatitude As New XElement("NorthLatitude", ds.Tables(0).Rows(RowNo).Item("AREA_NORTH_BOUND_LAT"))
                aou.Add(northLatitude)
                Dim westLongitude As New XElement("WestLongitude", ds.Tables(0).Rows(RowNo).Item("AREA_WEST_BOUND_LON"))
                aou.Add(westLongitude)
                Dim eastLongitude As New XElement("EastLongitude", ds.Tables(0).Rows(RowNo).Item("AREA_EAST_BOUND_LON"))
                aou.Add(eastLongitude)
                Dim isoA2Code As New XElement("IsoA2Code", ds.Tables(0).Rows(RowNo).Item("ISO_A2_CODE"))
                aou.Add(isoA2Code)
                Dim isoA3Code As New XElement("IsoA3Code", ds.Tables(0).Rows(RowNo).Item("ISO_A3_CODE"))
                aou.Add(isoA3Code)
                Dim isoNCode As New XElement("IsoNCode", ds.Tables(0).Rows(RowNo).Item("ISO_N_CODE"))
                aou.Add(isoNCode)

                'Main.MessageAdd("Looking for alias names " & vbCrLf)
                Dim aliasNames As New XElement("AliasNames")

                AOUCode = ds.Tables(0).Rows(RowNo).Item("AREA_CODE")
                'Main.MessageAdd("Area code: " & Str(AOUCode) & vbCrLf)

                expression = "[OBJECT_TABLE_NAME] = 'Area' AND [OBJECT_CODE] = " & Str(AOUCode)
                'Main.MessageAdd("Search expression: " & expression & vbCrLf)

                Dim result = ds.Tables(1).Select(expression)

                'Main.MessageAdd("Number of matches is: " & result.Count & vbCrLf)

                For Each item In result
                    'Main.MessageAdd(item.Item("ALIAS").ToString & vbCrLf)
                    Dim aliasName As New XElement("AliasName", item.Item("ALIAS").ToString)
                    aliasNames.Add(aliasName)
                Next


                aou.Add(aliasNames)

                aous.Add(aou)


            Next

            aouList.Add(aous)

            XDoc.Add(aouList)

        End Sub

        'Find the index of a record with the specified Author and Code
        Public Function FindIndex(ByVal Author As String, ByVal Code As Integer)
            If List.Count = 0 Then
                FindIndexResult = FindIndexResults.ListEmpty
                Return -1 'No record found.
            Else
                Dim Match = From Area In List Where Area.Author = Author And Area.Code = Code
                If Match.Count = 0 Then
                    'No record found.
                    FindIndexResult = FindIndexResults.NotFound
                    Return -1
                ElseIf Match.Count = 1 Then
                    'Single record found.
                    FindIndexResult = FindIndexResults.OK
                    Return List.FindIndex(Function(Area) Area.Author = Author And Area.Code = Code)
                ElseIf Match.Count > 1 Then
                    'Multiple records found.
                    FindIndexResult = FindIndexResults.ManyFound
                    Return List.FindIndex(Function(Area) Area.Author = Author And Area.Code = Code)
                End If
            End If
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.

        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'AreaOfUseList

    Public Class UnitOfMeasureList
        'Class used to store a list of Unit of Measure parameters.

        Public List As New List(Of UnitOfMeasure)

    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the UnitOfMeasure List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            If IsNothing(XDoc) Then
                Exit Sub
            End If

            CreationDate = XDoc.<UOMList>.<CreationDate>.Value
            LastEditDate = XDoc.<UOMList>.<LastEditDate>.Value
            Description = XDoc.<UOMList>.<Description>.Value

            Dim UOMs = From item In XDoc.<UOMList>.<UOM>

            List.Clear()
            For Each item In UOMs
                Dim NewUOM As New UnitOfMeasure
                NewUOM.Name = item.<Name>.Value
                NewUOM.Author = item.<Author>.Value
                NewUOM.Code = item.<Code>.Value
                NewUOM.Selected = item.<Selected>.Value
                Select Case item.<Type>.Value
                    Case "Angle"
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Angle
                    Case "Length"
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Length
                    Case "Scale"
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Scale
                    Case "Time"
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Time
                    Case Else
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Unknown
                End Select
                NewUOM.Comments = item.<Comments>.Value
                NewUOM.Deprecated = item.<Deprecated>.Value
                NewUOM.FactorB = item.<FactorB>.Value
                NewUOM.FactorC = item.<FactorC>.Value
                NewUOM.StandardUnitName = item.<StandardUnitName>.Value
                Dim aliasNames = From item2 In item.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewUOM.AddAlias(item3)
                Next
                List.Add(NewUOM)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If IsNothing(XDoc) Then
                Exit Sub
            End If

            CreationDate = XDoc.<UOMList>.<CreationDate>.Value
            LastEditDate = XDoc.<UOMList>.<LastEditDate>.Value
            Description = XDoc.<UOMList>.<Description>.Value

            Dim UOMs = From item In XDoc.<UOMList>.<UOM>

            List.Clear()
            For Each item In UOMs
                Dim NewUOM As New UnitOfMeasure
                NewUOM.Name = item.<Name>.Value
                NewUOM.Author = item.<Author>.Value
                NewUOM.Code = item.<Code>.Value
                NewUOM.Selected = item.<Selected>.Value
                Select Case item.<Type>.Value
                    Case "Angle"
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Angle
                    Case "Length"
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Length
                    Case "Scale"
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Scale
                    Case "Time"
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Time
                    Case Else
                        NewUOM.Type = UnitOfMeasure.UOMTypes.Unknown
                End Select
                NewUOM.Comments = item.<Comments>.Value
                NewUOM.Deprecated = item.<Deprecated>.Value
                NewUOM.FactorB = item.<FactorB>.Value
                NewUOM.FactorC = item.<FactorC>.Value
                NewUOM.StandardUnitName = item.<StandardUnitName>.Value
                Dim aliasNames = From item2 In item.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewUOM.AddAlias(item3)
                Next
                List.Add(NewUOM)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Unit Of Measure list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            ds.Clear()
            ds.Reset()

            da = New OleDb.OleDbDataAdapter("Select * From [Unit of Measure]", myConnection)
            TableName = "[Unit of Measure]"
            da.Fill(ds, TableName) 'Read the Unit of Measure table into ds

            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "[Alias]"
            da.Fill(ds, TableName) 'Read the Alias table into ds

            Dim StdUnitCode As Integer

            'The following query could be used to select a list of Unit of Measure alias names:
            Dim UOMCode As Integer
            Dim expression As String

            'Dim NRows As Integer = ds.Tables(0).Rows.Count
            Dim NRows As Integer = ds.Tables("[Unit of Measure]").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Unit Of Measure list"

            Dim RowNo As Integer

            For RowNo = 0 To NRows - 1 'loop through each row in the Unit of Measure table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewUom As New UnitOfMeasure
                NewUom.Name = ds.Tables("[Unit of Measure]").Rows(RowNo).Item("UNIT_OF_MEAS_NAME")
                NewUom.Author = "EPSG"
                NewUom.Code = ds.Tables("[Unit of Measure]").Rows(RowNo).Item("UOM_CODE")
                If IsDBNull(ds.Tables("[Unit of Measure]").Rows(RowNo).Item("REMARKS")) Then
                    NewUom.Comments = ""
                Else
                    NewUom.Comments = ds.Tables("[Unit of Measure]").Rows(RowNo).Item("REMARKS")
                End If
                NewUom.Deprecated = ds.Tables("[Unit of Measure]").Rows(RowNo).Item("DEPRECATED")

                Select Case ds.Tables(0).Rows(RowNo).Item("UNIT_OF_MEAS_TYPE")
                    Case "angle"
                        NewUom.Type = UnitOfMeasure.UOMTypes.Angle
                    Case "length"
                        NewUom.Type = UnitOfMeasure.UOMTypes.Length
                    Case "scale"
                        NewUom.Type = UnitOfMeasure.UOMTypes.Scale
                    Case "time"
                        NewUom.Type = UnitOfMeasure.UOMTypes.Time
                    Case Else
                        NewUom.Type = UnitOfMeasure.UOMTypes.Unknown
                End Select
                'NewUom.Type = ds.Tables(0).Rows(RowNo).Item("UNIT_OF_MEAS_TYPE")

                StdUnitCode = ds.Tables(0).Rows(RowNo).Item("TARGET_UOM_CODE")
                expression = "[UOM_Code] = " & Str(StdUnitCode)
                Dim StdUnitName = ds.Tables("[Unit of Measure]").Select(expression)
                If StdUnitName.Count > 0 Then
                    NewUom.StandardUnitName = StdUnitName(0).Item("UNIT_OF_MEAS_NAME").ToString
                End If
                If IsDBNull(ds.Tables("[Unit of Measure]").Rows(RowNo).Item("FACTOR_B")) Then
                    NewUom.FactorB = Double.NaN
                Else
                    NewUom.FactorB = ds.Tables("[Unit of Measure]").Rows(RowNo).Item("FACTOR_B")
                End If
                If IsDBNull(ds.Tables("[Unit of Measure]").Rows(RowNo).Item("FACTOR_C")) Then
                    NewUom.FactorC = Double.NaN
                Else
                    NewUom.FactorC = ds.Tables("[Unit of Measure]").Rows(RowNo).Item("FACTOR_C")
                End If


                UOMCode = ds.Tables("[Unit of Measure]").Rows(RowNo).Item("UOM_CODE")
                expression = "[OBJECT_TABLE_NAME] = 'Unit of Measure' AND [OBJECT_CODE] = " & Str(UOMCode)
                Dim result = ds.Tables("[Alias]").Select(expression)
                For Each item In result
                    NewUom.AddAlias(item.Item("ALIAS").ToString)
                Next
                List.Add(NewUom)
            Next

            myConnection.Close()

        End Sub

        'Function to return the list of Units of Measure as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Unit of Measure List File-->
                       <UOMList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <UOM>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Type><%= item.Type %></Type>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <FactorB><%= item.FactorB %></FactorB>
                                     <FactorC><%= item.FactorC %></FactorC>
                                     <StandardUnitName><%= item.StandardUnitName %></StandardUnitName>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                 </UOM>
                           %>
                       </UOMList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'UnitOfMeasureList -------------------------------------------------------------------------------------------------------------------------------------------------------------


    Public Class PrimeMeridianList '----------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store a list of Prime Meridian parameters.

        Public List As New List(Of PrimeMeridian)

    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the UnitOfMeasure List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<PMList>.<CreationDate>.Value
            LastEditDate = XDoc.<PMList>.<LastEditDate>.Value
            Description = XDoc.<PMList>.<Description>.Value

            Dim PMs = From item In XDoc.<PMList>.<PM>

            List.Clear()
            For Each item In PMs
                Dim NewPM As New PrimeMeridian
                NewPM.Name = item.<Name>.Value
                NewPM.Author = item.<Author>.Value
                NewPM.Code = item.<Code>.Value
                NewPM.Selected = item.<Selected>.Value
                NewPM.Comments = item.<Comments>.Value
                NewPM.Deprecated = item.<Deprecated>.Value
                Select Case item.<LongitudeUOM>.Value
                    Case "Degree"
                        NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Degree
                    Case "Gradian"
                        NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Gradian
                    Case "Sexagesimal_DMS"
                        NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Sexagesimal_DMS
                    Case Else
                        NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
                End Select
                NewPM.LongitudeFromGreenwich = item.<LongitudeFromGreenwich>.Value
                Dim aliasNames = From item2 In item.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewPM.AddAlias(item3)
                Next
                List.Add(NewPM)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            CreationDate = XDoc.<PMList>.<CreationDate>.Value
            LastEditDate = XDoc.<PMList>.<LastEditDate>.Value
            Description = XDoc.<PMList>.<Description>.Value

            Dim PMs = From item In XDoc.<PMList>.<PM>

            List.Clear()
            For Each item In PMs
                Dim NewPM As New PrimeMeridian
                NewPM.Name = item.<Name>.Value
                NewPM.Author = item.<Author>.Value
                NewPM.Code = item.<Code>.Value
                NewPM.Selected = item.<Selected>.Value
                NewPM.Comments = item.<Comments>.Value
                NewPM.Deprecated = item.<Deprecated>.Value
                Select Case item.<LongitudeUOM>.Value
                    Case "Degree"
                        NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Degree
                    Case "Gradian"
                        NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Gradian
                    Case "Sexagesimal_DMS"
                        NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Sexagesimal_DMS
                    Case Else
                        NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
                End Select
                NewPM.LongitudeFromGreenwich = item.<LongitudeFromGreenwich>.Value
                Dim aliasNames = From item2 In item.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewPM.AddAlias(item3)
                Next
                List.Add(NewPM)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Prime Meridian list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            ds.Clear()
            ds.Reset()

            da = New OleDb.OleDbDataAdapter("Select * From [Prime Meridian]", myConnection)
            TableName = "[Prime Meridian]"
            da.Fill(ds, TableName) 'Read the Unit of Measure table into ds

            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "[Alias]"
            da.Fill(ds, TableName) 'Read the Alias table into ds

            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            TableName = "[Unit of Measure]"
            da.Fill(ds, TableName) 'Read the Unit of Measure table into ds

            Dim PMCode As Integer
            Dim expression As String
            Dim UOMCode As Integer

            Dim NRows As Integer = ds.Tables("[Prime Meridian]").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Prime Meridian list"

            Dim RowNo As Integer

            RaiseEvent Message("Reading Prime Meridians from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the Unit of Measure table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewPM As New PrimeMeridian
                NewPM.Name = ds.Tables("[Prime Meridian]").Rows(RowNo).Item("PRIME_MERIDIAN_NAME")
                NewPM.Author = "EPSG"
                NewPM.Code = ds.Tables("[Prime Meridian]").Rows(RowNo).Item("PRIME_MERIDIAN_CODE")
                If IsDBNull(ds.Tables("[Prime Meridian]").Rows(RowNo).Item("REMARKS")) Then
                    NewPM.Comments = ""
                Else
                    NewPM.Comments = ds.Tables("[Prime Meridian]").Rows(RowNo).Item("REMARKS")
                End If
                NewPM.Deprecated = ds.Tables("[Prime Meridian]").Rows(RowNo).Item("DEPRECATED")
                NewPM.LongitudeFromGreenwich = ds.Tables("[Prime Meridian]").Rows(RowNo).Item("GREENWICH_LONGITUDE")

                UOMCode = ds.Tables("[Prime Meridian]").Rows(RowNo).Item("UOM_CODE")
                expression = "[UOM_CODE] = " & Str(UOMCode)
                Dim UomResult = ds.Tables(2).Select(expression)
                If UomResult.Count > 0 Then
                    Select Case UomResult(0).Item("UNIT_OF_MEAS_NAME").ToString
                        Case "degree"
                            NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Degree
                        Case "grad"
                            NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Gradian
                        Case "sexagesimal DMS"
                            NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Sexagesimal_DMS
                        Case Else
                            NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
                    End Select
                    'NewPM.LongitudeUOM = UomResult(0).Item("UNIT_OF_MEAS_NAME").ToString
                Else
                    'NewPM.LongitudeUOM = ""
                    NewPM.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
                End If

                PMCode = ds.Tables(0).Rows(RowNo).Item("PRIME_MERIDIAN_CODE")
                expression = "[OBJECT_TABLE_NAME] = 'Prime Meridian' AND [OBJECT_CODE] = " & Str(PMCode)
                Dim result = ds.Tables("[Alias]").Select(expression)
                For Each item In result
                    NewPM.AddAlias(item.Item("ALIAS").ToString)
                Next
                List.Add(NewPM)
            Next
            myConnection.Close()
        End Sub

        'Function to return the list of Prime Meridians as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Prime Meridian List File-->
                       <PMList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <PM>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <LongitudeUOM><%= item.LongitudeUOM %></LongitudeUOM>
                                     <LongitudeFromGreenwich><%= item.LongitudeFromGreenwich %></LongitudeFromGreenwich>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                 </PM>
                           %>
                       </PMList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'PrimeMeridianList -------------------------------------------------------------------------------------------------------------------------------------------------------------


    Public Class EllipsoidList '--------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store a list of Ellispoid parameters.

        Public List As New List(Of Ellipsoid)

    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Ellipsoid List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<EllipsoidList>.<CreationDate>.Value
            LastEditDate = XDoc.<EllipsoidList>.<LastEditDate>.Value
            Description = XDoc.<EllipsoidList>.<Description>.Value

            Dim Ellipsoids = From item In XDoc.<EllipsoidList>.<Ellipsoid>

            List.Clear()
            For Each item In Ellipsoids
                Dim NewEllipsoid As New Ellipsoid
                NewEllipsoid.Name = item.<Name>.Value
                NewEllipsoid.Author = item.<Author>.Value
                If item.<Code>.Value = Nothing Then
                    NewEllipsoid.Code = 0
                Else
                    NewEllipsoid.Code = item.<Code>.Value
                End If
                NewEllipsoid.Selected = item.<Selected>.Value
                NewEllipsoid.Comments = item.<Comments>.Value
                If item.<Deprecated>.Value = Nothing Then
                    NewEllipsoid.Deprecated = True
                Else
                    NewEllipsoid.Deprecated = item.<Deprecated>.Value
                End If
                Select Case item.<EllipsoidParameters>.Value
                    Case "SemiMajorAxis_InverseFlattening"
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_InverseFlattening
                    Case "SemiMajorAxis_SemiMinorAxis"
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_SemiMinorAxis
                    Case Else
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
                End Select
                If item.<SemiMajorAxis>.Value = Nothing Then
                    NewEllipsoid.SemiMajorAxis = Double.NaN
                Else
                    NewEllipsoid.SemiMajorAxis = item.<SemiMajorAxis>.Value
                End If
                If item.<InverseFlattening>.Value = Nothing Then
                    NewEllipsoid.InverseFlattening = Double.NaN
                Else
                    NewEllipsoid.InverseFlattening = item.<InverseFlattening>.Value
                End If
                If item.<SemiMinorAxis>.Value = Nothing Then
                    NewEllipsoid.SemiMinorAxis = Double.NaN
                Else 'Then
                    NewEllipsoid.SemiMinorAxis = item.<SemiMinorAxis>.Value
                End If
                If item.<Unit>.<Name>.Value = Nothing Then
                    NewEllipsoid.Unit.Name = ""
                Else
                    NewEllipsoid.Unit.Name = item.<Unit>.<Name>.Value
                End If
                If item.<Unit>.<Code>.Value = Nothing Then
                    NewEllipsoid.Unit.Code = 0
                Else
                    NewEllipsoid.Unit.Code = item.<Unit>.<Code>.Value
                End If
                If item.<Unit>.<Type>.Value = Nothing Then
                    NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Unknown
                Else
                    Select Case item.<Unit>.<Type>.Value
                        Case "Angle"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Angle
                        Case "Length"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Length
                        Case "Scale"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Scale
                        Case "Time"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Time
                        Case Else
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Unknown
                    End Select
                End If
                If item.<Unit>.<StandardUnitName>.Value = Nothing Then
                    NewEllipsoid.Unit.StandardUnitName = ""
                Else
                    NewEllipsoid.Unit.StandardUnitName = item.<Unit>.<StandardUnitName>.Value
                End If
                If item.<Unit>.<FactorB>.Value = Nothing Then
                    NewEllipsoid.Unit.FactorB = 0
                Else
                    NewEllipsoid.Unit.FactorB = item.<Unit>.<FactorB>.Value
                End If
                If item.<Unit>.<FactorC>.Value = Nothing Then
                    NewEllipsoid.Unit.FactorC = 0
                Else
                    NewEllipsoid.Unit.FactorC = item.<Unit>.<FactorC>.Value
                End If
                If item.<Unit>.<Comments>.Value = Nothing Then
                    NewEllipsoid.Unit.Comments = ""
                Else
                    NewEllipsoid.Unit.Comments = item.<Unit>.<Comments>.Value
                End If
                If item.<Unit>.<Deprecated>.Value = Nothing Then
                    NewEllipsoid.Unit.Deprecated = True
                Else
                    NewEllipsoid.Unit.Deprecated = item.<Unit>.<Deprecated>.Value
                End If

                Dim aliasNames = From item2 In item.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewEllipsoid.AddAlias(item3)
                Next
                List.Add(NewEllipsoid)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            CreationDate = XDoc.<EllipsoidList>.<CreationDate>.Value
            LastEditDate = XDoc.<EllipsoidList>.<LastEditDate>.Value
            Description = XDoc.<EllipsoidList>.<Description>.Value

            Dim Ellipsoids = From item In XDoc.<EllipsoidList>.<Ellipsoid>

            List.Clear()
            For Each item In Ellipsoids
                Dim NewEllipsoid As New Ellipsoid
                NewEllipsoid.Name = item.<Name>.Value
                NewEllipsoid.Author = item.<Author>.Value
                If item.<Code>.Value = Nothing Then
                    NewEllipsoid.Code = 0
                Else
                    NewEllipsoid.Code = item.<Code>.Value
                End If
                NewEllipsoid.Selected = item.<Selected>.Value
                NewEllipsoid.Comments = item.<Comments>.Value
                If item.<Deprecated>.Value = Nothing Then
                    NewEllipsoid.Deprecated = True
                Else
                    NewEllipsoid.Deprecated = item.<Deprecated>.Value
                End If
                Select Case item.<EllipsoidParameters>.Value
                    Case "SemiMajorAxis_InverseFlattening"
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_InverseFlattening
                    Case "SemiMajorAxis_SemiMinorAxis"
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_SemiMinorAxis
                    Case Else
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
                End Select
                If item.<SemiMajorAxis>.Value = Nothing Then
                    NewEllipsoid.SemiMajorAxis = Double.NaN
                Else
                    NewEllipsoid.SemiMajorAxis = item.<SemiMajorAxis>.Value
                End If
                If item.<InverseFlattening>.Value = Nothing Then
                    NewEllipsoid.InverseFlattening = Double.NaN
                Else
                    NewEllipsoid.InverseFlattening = item.<InverseFlattening>.Value
                End If
                If item.<SemiMinorAxis>.Value = Nothing Then
                    NewEllipsoid.SemiMinorAxis = Double.NaN
                Else 'Then
                    NewEllipsoid.SemiMinorAxis = item.<SemiMinorAxis>.Value
                End If
                If item.<Unit>.<Name>.Value = Nothing Then
                    NewEllipsoid.Unit.Name = ""
                Else
                    NewEllipsoid.Unit.Name = item.<Unit>.<Name>.Value
                End If
                If item.<Unit>.<Code>.Value = Nothing Then
                    NewEllipsoid.Unit.Code = 0
                Else
                    NewEllipsoid.Unit.Code = item.<Unit>.<Code>.Value
                End If
                If item.<Unit>.<Type>.Value = Nothing Then
                    NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Unknown
                Else
                    Select Case item.<Unit>.<Type>.Value
                        Case "Angle"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Angle
                        Case "Length"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Length
                        Case "Scale"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Scale
                        Case "Time"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Time
                        Case Else
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Unknown
                    End Select
                End If
                If item.<Unit>.<StandardUnitName>.Value = Nothing Then
                    NewEllipsoid.Unit.StandardUnitName = ""
                Else
                    NewEllipsoid.Unit.StandardUnitName = item.<Unit>.<StandardUnitName>.Value
                End If
                If item.<Unit>.<FactorB>.Value = Nothing Then
                    NewEllipsoid.Unit.FactorB = 0
                Else
                    NewEllipsoid.Unit.FactorB = item.<Unit>.<FactorB>.Value
                End If
                If item.<Unit>.<FactorC>.Value = Nothing Then
                    NewEllipsoid.Unit.FactorC = 0
                Else
                    NewEllipsoid.Unit.FactorC = item.<Unit>.<FactorC>.Value
                End If
                If item.<Unit>.<Comments>.Value = Nothing Then
                    NewEllipsoid.Unit.Comments = ""
                Else
                    NewEllipsoid.Unit.Comments = item.<Unit>.<Comments>.Value
                End If
                If item.<Unit>.<Deprecated>.Value = Nothing Then
                    NewEllipsoid.Unit.Deprecated = True
                Else
                    NewEllipsoid.Unit.Deprecated = item.<Unit>.<Deprecated>.Value
                End If


                Dim aliasNames = From item2 In item.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewEllipsoid.AddAlias(item3)
                Next
                List.Add(NewEllipsoid)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Ellipsoid list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            ds.Clear()
            ds.Reset()

            da = New OleDb.OleDbDataAdapter("Select * From [Ellipsoid]", myConnection)
            TableName = "[Ellipsoid]"
            da.Fill(ds, TableName) 'Read the Unit of Measure table into ds
            'ELLIPSOID_CODE ELLIPSOID_NAME SEMI_MAJOR_AXIS UOM_CODE INV_FLATTENING SEMI_MINOR_AXIS ELLIPSOID_SHAPE REMARKS DEPRECATED

            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "[Alias]"
            da.Fill(ds, TableName) 'Read the Alias table into ds

            'Read the Unit of Measure table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            TableName = "Unit"
            da.Fill(ds, TableName)
            'UOM_CODE UNIT_OF_MEAS_NAME UNIT_OF_MEAS_TYPE TARGET_UOM_CODE FACTOR_B FACTOR_C REMARKS DEPRECATED


            Dim EllipsoidCode As Integer
            'Dim UnitCode As Integer
            Dim TargetUomCode As Integer 'The standard unit code
            Dim expression As String

            Dim NRows As Integer = ds.Tables("[Ellipsoid]").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Ellipsoid list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            'Dim EllipsoidParametersString As String
            RaiseEvent Message("Reading Ellipsoids from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the Ellipsoid table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewEllipsoid As New Ellipsoid
                NewEllipsoid.Name = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("ELLIPSOID_NAME")
                NewEllipsoid.Author = "EPSG"
                NewEllipsoid.Code = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("ELLIPSOID_CODE")
                If IsDBNull(ds.Tables("[Ellipsoid]").Rows(RowNo).Item("REMARKS")) Then
                    NewEllipsoid.Comments = ""
                Else
                    NewEllipsoid.Comments = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("REMARKS")
                End If

                NewEllipsoid.Deprecated = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("DEPRECATED")

                If IsDBNull(ds.Tables("[Ellipsoid]").Rows(RowNo).Item("INV_FLATTENING")) Then
                    If IsDBNull(ds.Tables("[Ellipsoid]").Rows(RowNo).Item("SEMI_MINOR_AXIS")) Then
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
                    End If
                    NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_SemiMinorAxis
                Else
                    If IsDBNull(ds.Tables("[Ellipsoid]").Rows(RowNo).Item("SEMI_MINOR_AXIS")) Then
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_InverseFlattening
                    Else
                        NewEllipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
                    End If
                End If

                If IsDBNull(ds.Tables("[Ellipsoid]").Rows(RowNo).Item("SEMI_MAJOR_AXIS")) Then
                    NewEllipsoid.SemiMajorAxis = Double.NaN
                Else
                    NewEllipsoid.SemiMajorAxis = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("SEMI_MAJOR_AXIS")
                End If
                If IsDBNull(ds.Tables("[Ellipsoid]").Rows(RowNo).Item("INV_FLATTENING")) Then
                    NewEllipsoid.InverseFlattening = Double.NaN
                Else
                    NewEllipsoid.InverseFlattening = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("INV_FLATTENING")
                End If
                If IsDBNull(ds.Tables("[Ellipsoid]").Rows(RowNo).Item("SEMI_MINOR_AXIS")) Then
                    NewEllipsoid.SemiMinorAxis = Double.NaN
                Else
                    NewEllipsoid.SemiMinorAxis = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("SEMI_MINOR_AXIS")
                End If

                NewEllipsoid.Unit.Code = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("UOM_CODE")
                expression = "[UOM_CODE] = " & Str(NewEllipsoid.Unit.Code)
                Dim unitResult = ds.Tables("Unit").Select(expression)
                If unitResult.Count = 0 Then
                    RaiseEvent ErrorMessage("Missing axis unit. Ellipsoid name: " & ds.Tables("[Ellipsoid]").Rows(RowNo).Item("ELLIPSOID_NAME") & vbCrLf)
                ElseIf unitResult.Count > 1 Then
                    RaiseEvent ErrorMessage("Greater than one axis unit. Ellipsoid name: " & ds.Tables("[Ellipsoid]").Rows(RowNo).Item("ELLIPSOID_NAME") & vbCrLf)
                Else

                    NewEllipsoid.Unit.Name = unitResult(0).Item("UNIT_OF_MEAS_NAME").ToString
                    Select Case unitResult(0).Item("UNIT_OF_MEAS_TYPE").ToString
                        Case "Angle"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Angle
                        Case "Length"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Length
                        Case "Scale"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Scale
                        Case "Time"
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Time
                        Case Else
                            NewEllipsoid.Unit.Type = UnitOfMeasure.UOMTypes.Unknown
                    End Select

                    TargetUomCode = unitResult(0).Item("TARGET_UOM_CODE")
                    expression = "[UOM_CODE] = " & Str(TargetUomCode)
                    Dim targetUnitResult = ds.Tables("Unit").Select(expression)
                    If targetUnitResult.Count = 0 Then
                        RaiseEvent ErrorMessage("Missing axis target unit. Coordinate System name: " & ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME") & vbCrLf)
                    ElseIf targetUnitResult.Count > 1 Then
                        RaiseEvent ErrorMessage("Greater than one axis target unit. Coordinate System name: " & ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME") & vbCrLf)
                    Else
                        NewEllipsoid.Unit.StandardUnitName = targetUnitResult(0).Item("UNIT_OF_MEAS_NAME").ToString
                        If IsDBNull(unitResult(0).Item("FACTOR_B")) Then
                            NewEllipsoid.Unit.FactorB = 1
                        Else
                            NewEllipsoid.Unit.FactorB = unitResult(0).Item("FACTOR_B")
                        End If
                        If IsDBNull(unitResult(0).Item("FACTOR_C")) Then
                            NewEllipsoid.Unit.FactorC = 1
                        Else
                            NewEllipsoid.Unit.FactorC = unitResult(0).Item("FACTOR_C")
                        End If


                        If IsDBNull(unitResult(0).Item("REMARKS")) Then
                            NewEllipsoid.Unit.Comments = ""
                        Else
                            NewEllipsoid.Unit.Comments = unitResult(0).Item("REMARKS")
                        End If

                        NewEllipsoid.Unit.Deprecated = unitResult(0).Item("DEPRECATED")
                    End If
                End If

                EllipsoidCode = ds.Tables("[Ellipsoid]").Rows(RowNo).Item("ELLIPSOID_CODE")
                expression = "[OBJECT_TABLE_NAME] = 'Ellipsoid' AND [OBJECT_CODE] = " & Str(EllipsoidCode)
                Dim result = ds.Tables("[Alias]").Select(expression)
                For Each item In result
                    NewEllipsoid.AddAlias(item.Item("ALIAS").ToString)
                Next
                List.Add(NewEllipsoid)
            Next
            myConnection.Close()
        End Sub

        'Function to return the list of Prime Meridians as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Prime Meridian List File-->
                       <EllipsoidList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <Ellipsoid>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <EllipsoidParameters><%= item.EllipsoidParameters %></EllipsoidParameters>
                                     <SemiMajorAxis><%= item.SemiMajorAxis %></SemiMajorAxis>
                                     <InverseFlattening><%= item.InverseFlattening %></InverseFlattening>
                                     <SemiMinorAxis><%= item.SemiMinorAxis %></SemiMinorAxis>
                                     <Unit>
                                         <Name><%= item.Unit.Name %></Name>
                                         <Code><%= item.Unit.Code %></Code>
                                         <Type><%= item.Unit.Type %></Type>
                                         <StandardUnitName><%= item.Unit.StandardUnitName %></StandardUnitName>
                                         <FactorB><%= item.Unit.FactorB %></FactorB>
                                         <FactorC><%= item.Unit.FactorC %></FactorC>
                                         <Comments><%= item.Unit.Comments %></Comments>
                                         <Deprecated><%= item.Unit.Deprecated %></Deprecated>
                                     </Unit>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                 </Ellipsoid>
                           %>
                       </EllipsoidList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'EllispoidList -----------------------------------------------------------------------------------------------------------------------------------------------------------------


    Public Class ProjectionList '-------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store Projection parameters.

        Public List As New List(Of Projection)

    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Projection List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<ProjectionList>.<CreationDate>.Value
            LastEditDate = XDoc.<ProjectionList>.<LastEditDate>.Value
            Description = XDoc.<ProjectionList>.<Description>.Value

            Dim Projections = From item In XDoc.<ProjectionList>.<Projection>

            List.Clear()
            For Each projectionItem In Projections
                Dim NewProjection As New Projection
                NewProjection.Name = projectionItem.<Name>.Value
                NewProjection.Author = projectionItem.<Author>.Value
                If projectionItem.<Code>.Value = Nothing Then
                    NewProjection.Code = 0
                Else
                    NewProjection.Code = projectionItem.<Code>.Value
                End If
                NewProjection.Selected = projectionItem.<Selected>.Value
                NewProjection.Comments = projectionItem.<Comments>.Value
                If projectionItem.<Scope>.Value = Nothing Then
                    NewProjection.Scope = ""
                Else
                    NewProjection.Scope = projectionItem.<Scope>.Value
                End If
                If projectionItem.<Deprecated>.Value = Nothing Then
                    NewProjection.Deprecated = True
                Else
                    NewProjection.Deprecated = projectionItem.<Deprecated>.Value
                End If
                NewProjection.Method.Name = projectionItem.<ProjectionMethod>.<Name>.Value
                NewProjection.Method.Author = projectionItem.<ProjectionMethod>.<Author>.Value
                NewProjection.Method.Code = projectionItem.<ProjectionMethod>.<Code>.Value
                NewProjection.Area.Name = projectionItem.<AreaOfUse>.<Name>.Value
                NewProjection.Area.Author = projectionItem.<AreaOfUse>.<Author>.Value
                NewProjection.Area.Code = projectionItem.<AreaOfUse>.<Code>.Value

                Dim parameters = From item In projectionItem.<Parameters>.<Parameter>

                For Each parameterItem In parameters
                    Dim NewParameter As New ValueSummary
                    NewParameter.Name = parameterItem.<Name>.Value
                    NewParameter.Value = parameterItem.<Value>.Value
                    NewParameter.Unit.Name = parameterItem.<UnitOfMeasure>.<Name>.Value
                    NewParameter.Unit.Author = parameterItem.<UnitOfMeasure>.<Author>.Value
                    NewParameter.Unit.Code = parameterItem.<UnitOfMeasure>.<Code>.Value
                    NewProjection.ParameterValue.Add(NewParameter)
                Next


                Dim aliasNames = From item2 In projectionItem.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewProjection.AddAlias(item3)
                Next
                List.Add(NewProjection)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            CreationDate = XDoc.<ProjectionList>.<CreationDate>.Value
            LastEditDate = XDoc.<ProjectionList>.<LastEditDate>.Value
            Description = XDoc.<ProjectionList>.<Description>.Value

            Dim Projections = From item In XDoc.<ProjectionList>.<Projection>

            List.Clear()
            For Each projectionItem In Projections
                Dim NewProjection As New Projection
                NewProjection.Name = projectionItem.<Name>.Value
                NewProjection.Author = projectionItem.<Author>.Value
                If projectionItem.<Code>.Value = Nothing Then
                    NewProjection.Code = 0
                Else
                    NewProjection.Code = projectionItem.<Code>.Value
                End If
                NewProjection.Selected = projectionItem.<Selected>.Value
                NewProjection.Comments = projectionItem.<Comments>.Value
                If projectionItem.<Scope>.Value = Nothing Then
                    NewProjection.Scope = ""
                Else
                    NewProjection.Scope = projectionItem.<Scope>.Value
                End If
                If projectionItem.<Deprecated>.Value = Nothing Then
                    NewProjection.Deprecated = True
                Else
                    NewProjection.Deprecated = projectionItem.<Deprecated>.Value
                End If
                NewProjection.Method.Name = projectionItem.<ProjectionMethod>.<Name>.Value
                NewProjection.Method.Author = projectionItem.<ProjectionMethod>.<Author>.Value
                NewProjection.Method.Code = projectionItem.<ProjectionMethod>.<Code>.Value

                NewProjection.Area.Name = projectionItem.<AreaOfUse>.<Name>.Value
                NewProjection.Area.Author = projectionItem.<AreaOfUse>.<Author>.Value
                NewProjection.Area.Code = projectionItem.<AreaOfUse>.<Code>.Value

                Dim parameters = From item In projectionItem.<Parameters>.<Parameter>

                For Each parameterItem In parameters
                    Dim NewParameter As New ValueSummary
                    NewParameter.Name = parameterItem.<Name>.Value
                    NewParameter.Value = parameterItem.<Value>.Value
                    NewParameter.Unit.Name = parameterItem.<UnitOfMeasure>.<Name>.Value
                    NewParameter.Unit.Author = parameterItem.<UnitOfMeasure>.<Author>.Value
                    NewParameter.Unit.Code = parameterItem.<UnitOfMeasure>.<Code>.Value
                    NewProjection.ParameterValue.Add(NewParameter)
                Next


                Dim aliasNames = From item2 In projectionItem.<AliasNames>.<AliasName>
                For Each item3 In aliasNames
                    NewProjection.AddAlias(item3)
                Next
                List.Add(NewProjection)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Ellipsoid list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Coordinate_Operation table into dataset ds - Only read the "conversion" records.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate_Operation] Where [COORD_OP_TYPE] = 'conversion'", myConnection)
            TableName = "[CoordOp]"
            da.Fill(ds, TableName) 'Read the Unit of Measure table into ds

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "[Alias]"
            da.Fill(ds, TableName) 'Read the Alias table into ds

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "[Area]"
            da.Fill(ds, TableName)

            'Read the Coordinate Operation Method table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Method]", myConnection)
            TableName = "[CoordOpMethod]"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, FORMULA, EXAMPLE, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            'Read the Coordinate Operation Parameter table into dataset ds:
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter]", myConnection)
            TableName = "[CoordOpParams]"
            da.Fill(ds, TableName)
            'Table fields: PARAMETER_CODE, PARAMETER_NAME, DESCRIPTION, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            'Read the Coordinate Operation Parameter Value table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter Value]", myConnection)
            TableName = "[CoordOpValues]"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_CODE, COORD_OP_METHOD_CODE, PARAMETER_CODE, PARAMETER_VALUE, PARAM_VALUE_FILE_REF, UOM_CODE

            'Read the Coordinate Operation Parameter Usage table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter Usage]", myConnection)
            TableName = "[CoordOpUsage]"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_METHOD_CODE, PARAMETER_CODE, SORT_ORDER, PARAM_SIGN_REVERSAL

            'Read the Unit of Measure table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            TableName = "[UnitOfMeasure]"
            da.Fill(ds, TableName)
            'Table fields: UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED


            'Dim CoordOpCode As Integer
            'Dim CoordOpMethodCode As Integer
            'Dim AreaOfUseCode As Integer

            Dim expression As String

            'Projection parameter variables:
            Dim NParams As Integer 'The number of parameters used to define the projection.
            Dim ParamNo As Integer 'The current parameter number.
            Dim ParamCode As Integer 'The parameter code of the current parameter
            Dim UomCode As Integer 'The unit of measure code corresponding to the parameter
            Dim TargetUomCode As Integer 'The standard unit code

            Dim NRows As Integer = ds.Tables("[CoordOp]").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Projection list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Projections from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordOp table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewProjection As New Projection
                NewProjection.Name = ds.Tables("[CoordOp]").Rows(RowNo).Item("COORD_OP_NAME")
                NewProjection.Author = "EPSG"
                NewProjection.Code = ds.Tables("[CoordOp]").Rows(RowNo).Item("COORD_OP_CODE")
                If IsDBNull(ds.Tables("[CoordOp]").Rows(RowNo).Item("REMARKS")) Then
                    NewProjection.Comments = ""
                Else
                    NewProjection.Comments = ds.Tables("[CoordOp]").Rows(RowNo).Item("REMARKS")
                End If

                NewProjection.Scope = ds.Tables("[CoordOp]").Rows(RowNo).Item("COORD_OP_SCOPE")
                NewProjection.Deprecated = ds.Tables("[CoordOp]").Rows(RowNo).Item("DEPRECATED")

                NewProjection.Method.Code = ds.Tables("[CoordOp]").Rows(RowNo).Item("COORD_OP_METHOD_CODE")
                expression = "[COORD_OP_METHOD_CODE] = " & Str(NewProjection.Method.Code)
                Dim CoordOpMethodParameters = ds.Tables("[CoordOpMethod]").Select(expression)
                If CoordOpMethodParameters.Count > 0 Then
                    NewProjection.Method.Name = CoordOpMethodParameters(0).Item("COORD_OP_METHOD_NAME").ToString
                    NewProjection.Method.Author = "EPSG"
                End If

                NewProjection.Area.Code = ds.Tables("[CoordOp]").Rows(RowNo).Item("AREA_OF_USE_CODE")
                expression = "[AREA_CODE] = " & Str(NewProjection.Area.Code)
                Dim areaOfUseParameters = ds.Tables("[Area]").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewProjection.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                    NewProjection.Area.Author = "EPSG"
                End If

                expression = "[COORD_OP_METHOD_CODE] = " & Str(NewProjection.Method.Code)
                Dim CoordOpUsage = ds.Tables("[CoordOpUsage]").Select(expression)
                NParams = CoordOpUsage.Count
                For ParamNo = 0 To NParams - 1 'Process each Projection Parameter.
                    Dim NewValueSummary As New ValueSummary
                    ParamCode = CoordOpUsage(ParamNo).Item("PARAMETER_CODE")

                    'Get the Projection Parameter Details:
                    expression = "[PARAMETER_CODE] = " & Str(ParamCode)
                    Dim parameterDetails = ds.Tables("[CoordOpParams]").Select(expression)
                    'Table fields: PARAMETER_CODE, PARAMETER_NAME, DESCRIPTION, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

                    'Get the Projection Parameter Value:
                    expression = "[COORD_OP_CODE] = " & Str(NewProjection.Code) & " And [PARAMETER_CODE] = " & Str(ParamCode)
                    Dim parameterValues = ds.Tables("[CoordOpValues]").Select(expression)
                    'Table fields: COORD_OP_CODE, COORD_OP_METHOD_CODE, PARAMETER_CODE, PARAMETER_VALUE, PARAM_VALUE_FILE_REF, UOM_CODE

                    UomCode = parameterValues(0).Item("UOM_CODE")

                    'Get the Unit Of Measure Details:
                    expression = "[UOM_CODE] = " & Str(UomCode)
                    Dim parameterUom = ds.Tables("[UnitOfMeasure]").Select(expression)
                    'Table fields: UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

                    NewValueSummary.Name = parameterDetails(0).Item("PARAMETER_NAME")
                    NewValueSummary.Value = parameterValues(0).Item("PARAMETER_VALUE")
                    NewValueSummary.Unit.Name = parameterUom(0).Item("UNIT_OF_MEAS_NAME")
                    NewValueSummary.Unit.Author = "EPSG"
                    NewValueSummary.Unit.Code = Str(UomCode)

                    NewProjection.ParameterValue.Add(NewValueSummary)
                Next
                List.Add(NewProjection)
            Next

            myConnection.Close()
        End Sub

        'Function to return the list of Projections as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Prime Meridian List File-->
                       <ProjectionList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <Projection>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <ProjectionMethod>
                                         <Name><%= item.Method.Name %></Name>
                                         <Author><%= item.Method.Author %></Author>
                                         <Code><%= item.Method.Code %></Code>
                                     </ProjectionMethod>
                                     <Parameters>
                                         <%= From paramItem In item.ParameterValue
                                             Select _
                                             <Parameter>
                                                 <Name><%= paramItem.Name %></Name>
                                                 <Value><%= paramItem.Value %></Value>
                                                 <UnitOfMeasure>
                                                     <Name><%= paramItem.Unit.Name %></Name>
                                                     <Author><%= paramItem.Unit.Author %></Author>
                                                     <Code><%= paramItem.Unit.Code %></Code>
                                                 </UnitOfMeasure>
                                             </Parameter>
                                         %>
                                     </Parameters>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                 </Projection>
                           %>
                       </ProjectionList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'ProjectionList ----------------------------------------------------------------------------------------------------------------------------------------------------------------

#End Region 'Lists of Parameters -------------------------------------------------------------------------------------------------------------------------------------------------------------


    Public Class CoordOpMethodList '----------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store Coordinate Operation Method parameters.

        'Public List As New List(Of CoordinateOperationMethod)
        Public List As New List(Of CoordOpMethod)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the CoordOpMethod List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<CoordOpMethodList>.<CreationDate>.Value
            LastEditDate = XDoc.<CoordOpMethodList>.<LastEditDate>.Value
            Description = XDoc.<CoordOpMethodList>.<Description>.Value

            Dim Methods = From item In XDoc.<CoordOpMethodList>.<CoordOpMethod>

            List.Clear()
            For Each methodItem In Methods
                Dim NewMethod As New CoordOpMethod
                NewMethod.Name = methodItem.<Name>.Value
                NewMethod.Author = methodItem.<Author>.Value
                If methodItem.<Code>.Value = Nothing Then
                    NewMethod.Code = 0
                Else
                    NewMethod.Code = methodItem.<Code>.Value
                End If
                NewMethod.ReverseOp = methodItem.<ReverseOp>.Value
                NewMethod.Formula = methodItem.<Formula>.Value
                NewMethod.Example = methodItem.<Example>.Value
                NewMethod.Selected = methodItem.<Selected>.Value
                NewMethod.Comments = methodItem.<Comments>.Value

                If methodItem.<Deprecated>.Value = Nothing Then
                    NewMethod.Deprecated = True
                Else
                    NewMethod.Deprecated = methodItem.<Deprecated>.Value
                End If

                Dim parameters = From item In methodItem.<ParameterList>.<Parameter>

                For Each parameterItem In parameters
                    Dim NewParameter As New ParameterSummary
                    NewParameter.Name = parameterItem.<Name>.Value
                    NewParameter.Description = parameterItem.<Description>.Value
                    NewParameter.Order = parameterItem.<Order>.Value
                    NewParameter.SignReversal = parameterItem.<SignReversal>.Value
                    NewMethod.Parameter.Add(NewParameter)
                Next

                List.Add(NewMethod)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            CreationDate = XDoc.<CoordOpMethodList>.<CreationDate>.Value
            LastEditDate = XDoc.<CoordOpMethodList>.<LastEditDate>.Value
            Description = XDoc.<CoordOpMethodList>.<Description>.Value

            Dim Methods = From item In XDoc.<CoordOpMethodList>.<CoordOpMethod>

            List.Clear()
            For Each methodItem In Methods
                Dim NewMethod As New CoordOpMethod
                NewMethod.Name = methodItem.<Name>.Value
                NewMethod.Author = methodItem.<Author>.Value
                If methodItem.<Code>.Value = Nothing Then
                    NewMethod.Code = 0
                Else
                    NewMethod.Code = methodItem.<Code>.Value
                End If
                NewMethod.ReverseOp = methodItem.<ReverseOp>.Value
                NewMethod.Formula = methodItem.<Formula>.Value
                NewMethod.Example = methodItem.<Example>.Value
                NewMethod.Selected = methodItem.<Selected>.Value
                NewMethod.Comments = methodItem.<Comments>.Value

                If methodItem.<Deprecated>.Value = Nothing Then
                    NewMethod.Deprecated = True
                Else
                    NewMethod.Deprecated = methodItem.<Deprecated>.Value
                End If

                Dim parameters = From item In methodItem.<ParameterList>.<Parameter>

                For Each parameterItem In parameters
                    Dim NewParameter As New ParameterSummary
                    NewParameter.Name = parameterItem.<Name>.Value
                    NewParameter.Description = parameterItem.<Description>.Value
                    NewParameter.Order = parameterItem.<Order>.Value
                    NewParameter.SignReversal = parameterItem.<SignReversal>.Value
                    NewMethod.Parameter.Add(NewParameter)
                Next

                List.Add(NewMethod)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Coordinate Operation Method list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString
            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Coordinate Operation Method table into dataset ds
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate_Operation Method]", myConnection)
            TableName = "CoordOpMethod"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, FORMULA, EXAMPLE, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName) 'Read the Alias table into ds

            'Read the Coordinate Operation Parameter Usage table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter Usage]", myConnection)
            TableName = "CoordOpUsage"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_METHOD_CODE, PARAMETER_CODE, SORT_ORDER, PARAM_SIGN_REVERSAL

            'Read the Coordinate Operation Parameter table into dataset ds:
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter]", myConnection)
            TableName = "CoordOpParams"
            da.Fill(ds, TableName)
            'Table fields: PARAMETER_CODE, PARAMETER_NAME, DESCRIPTION, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            Dim expression As String

            ' variables:
            Dim NParams As Integer 'The number of parameters used to define the projection.
            Dim ParamNo As Integer 'The current parameter number.
            Dim ParamCode As Integer 'The parameter code of the current parameter
            'Dim UomCode As Integer 'The unit of measure code corresponding to the parameter
            'Dim TargetUomCode As Integer 'The standard unit code

            Dim NRows As Integer = ds.Tables("CoordOpMethod").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Coordinate Operation Method list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Coordinate Operation Methods from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordOpMethod table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewCoordOpMethod As New CoordOpMethod
                NewCoordOpMethod.Name = ds.Tables("CoordOpMethod").Rows(RowNo).Item("COORD_OP_METHOD_NAME") '.ToString ???
                NewCoordOpMethod.Author = "EPSG"
                NewCoordOpMethod.Code = ds.Tables("CoordOpMethod").Rows(RowNo).Item("COORD_OP_METHOD_CODE")
                If IsDBNull(ds.Tables("CoordOpMethod").Rows(RowNo).Item("REMARKS")) Then
                    NewCoordOpMethod.Comments = ""
                Else
                    NewCoordOpMethod.Comments = ds.Tables("CoordOpMethod").Rows(RowNo).Item("REMARKS")
                End If
                Select Case ds.Tables("CoordOpMethod").Rows(RowNo).Item("REVERSE_OP").ToString
                    Case "True"
                        NewCoordOpMethod.ReverseOp = True
                    Case "False"
                        NewCoordOpMethod.ReverseOp = False
                    Case Else
                        NewCoordOpMethod.ReverseOp = False
                        RaiseEvent ErrorMessage("Unrecognised ReverseOp text: " & ds.Tables("CoordOpMethod").Rows(RowNo).Item("REVERSE_OP").ToString & " of Method: " & NewCoordOpMethod.Name & vbCrLf)
                End Select
                NewCoordOpMethod.Formula = ds.Tables("CoordOpMethod").Rows(RowNo).Item("FORMULA").ToString
                NewCoordOpMethod.Example = ds.Tables("CoordOpMethod").Rows(RowNo).Item("EXAMPLE").ToString
                NewCoordOpMethod.Deprecated = ds.Tables("CoordOpMethod").Rows(RowNo).Item("DEPRECATED")

                'Get the list of parameters for this Coordinate Operation Method:
                expression = "[COORD_OP_METHOD_CODE] = " & Str(NewCoordOpMethod.Code)
                Dim CoordOpMethodUsage = ds.Tables("CoordOpUsage").Select(expression)
                'Table fields: COORD_OP_METHOD_CODE, PARAMETER_CODE, SORT_ORDER, PARAM_SIGN_REVERSAL

                NParams = CoordOpMethodUsage.Count
                If NParams > 0 Then
                    For ParamNo = 0 To NParams - 1
                        ParamCode = CoordOpMethodUsage(ParamNo).Item("PARAMETER_CODE")
                        'Get the Coordinate Operation Method Parameter details:
                        expression = "[PARAMETER_CODE] = " & Str(ParamCode)
                        Dim parameterDetails = ds.Tables("CoordOpParams").Select(expression)
                        'Table fields: PARAMETER_CODE, PARAMETER_NAME, DESCRIPTION, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED
                        If parameterDetails.Count > 0 Then
                            Dim NewParameter As New ParameterSummary
                            NewParameter.Name = parameterDetails(0).Item("PARAMETER_NAME")
                            NewParameter.Description = parameterDetails(0).Item("DESCRIPTION")
                            NewParameter.Order = CoordOpMethodUsage(ParamNo).Item("SORT_ORDER")
                            If IsDBNull(CoordOpMethodUsage(ParamNo).Item("PARAM_SIGN_REVERSAL")) Then
                                NewParameter.SignReversal = False
                            Else
                                Select Case CoordOpMethodUsage(ParamNo).Item("PARAM_SIGN_REVERSAL")
                                    Case "Yes"
                                        NewParameter.SignReversal = True
                                    Case "No"
                                        NewParameter.SignReversal = False
                                    Case Else
                                        NewParameter.SignReversal = False
                                        RaiseEvent ErrorMessage("Unrecognised SignReversal text: " & CoordOpMethodUsage(ParamNo).Item("PARAM_SIGN_REVERSAL") & " in Parameter: " & NewParameter.Name & " of Method: " & NewCoordOpMethod.Name & vbCrLf)
                                End Select
                                'NewParameter.SignReversal = CoordOpMethodUsage(ParamNo).Item("PARAM_SIGN_REVERSAL")
                            End If

                            NewCoordOpMethod.Parameter.Add(NewParameter)
                        End If
                    Next
                End If
                List.Add(NewCoordOpMethod)
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Coordinate Operation Methods as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Coordinate Operation Method List File-->
                       <CoordOpMethodList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <CoordOpMethod>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <ReverseOp><%= item.ReverseOp %></ReverseOp>
                                     <Formula><%= item.Formula %></Formula>
                                     <Example><%= item.Example %></Example>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <ParameterList>
                                         <%= From paramItem In item.Parameter
                                             Select _
                                             <Parameter>
                                                 <Name><%= paramItem.Name %></Name>
                                                 <Description><%= paramItem.Description %></Description>
                                                 <Order><%= paramItem.Order %></Order>
                                                 <SignReversal><%= paramItem.SignReversal %></SignReversal>
                                             </Parameter>
                                         %>
                                     </ParameterList>
                                 </CoordOpMethod>
                           %>
                       </CoordOpMethodList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'CoordOpMethodList -------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class CoordRefSystemList '---------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store Coordinate Reference System parameters.

        Public List As New List(Of CoordinateReferenceSystemSummary)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the CoordRefSystem List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<CoordinateReferenceSystemList>.<CreationDate>.Value
            LastEditDate = XDoc.<CoordinateReferenceSystemList>.<LastEditDate>.Value
            Description = XDoc.<CoordinateReferenceSystemList>.<Description>.Value

            Dim Systems = From item In XDoc.<CoordinateReferenceSystemList>.<CoordinateReferenceSystem>

            List.Clear()
            For Each systemItem In Systems
                Dim NewSystem As New CoordinateReferenceSystemSummary
                NewSystem.Name = systemItem.<Name>.Value
                NewSystem.Author = systemItem.<Author>.Value
                If systemItem.<Code>.Value = Nothing Then
                    NewSystem.Code = 0
                Else
                    NewSystem.Code = systemItem.<Code>.Value
                End If
                Select Case systemItem.<Type>.Value
                    Case "Compound" '"compound"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Compound
                        NewSystem.Type = CrsTypes.Compound
                    Case "Engineering" '"engineering"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Engineering
                        NewSystem.Type = CrsTypes.Engineering
                    Case "Geocentric" '"geocentric"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geocentric
                        NewSystem.Type = CrsTypes.Geocentric
                    Case "Geographic2D" '"geographic 2D"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic2D
                        NewSystem.Type = CrsTypes.Geographic2D
                    Case "Geographic3D" '"geographic 3D"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic3D
                        NewSystem.Type = CrsTypes.Geographic3D
                    Case "Projected" '"projected"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Projected
                        NewSystem.Type = CrsTypes.Projected
                    Case "Vertical" '"vertical"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Vertical
                        NewSystem.Type = CrsTypes.Vertical
                    Case Else
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Unknown
                        NewSystem.Type = CrsTypes.Unknown
                End Select
                NewSystem.Scope = systemItem.<Scope>.Value
                NewSystem.Comments = systemItem.<Comments>.Value
                NewSystem.Deprecated = systemItem.<Deprecated>.Value

                Dim aliasNames = From item In systemItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewSystem.AddAlias(nameItem)
                Next

                'Read Area of Use information: --------------------------------------------------------------
                NewSystem.Area.Name = systemItem.<AreaOfUse>.<Name>.Value
                NewSystem.Area.Author = systemItem.<AreaOfUse>.<Author>.Value
                NewSystem.Area.Code = systemItem.<AreaOfUse>.<Code>.Value
                NewSystem.Area.Type = ""

                'Read Coordinate System information: --------------------------------------------------------------
                NewSystem.CoordinateSystem.Name = systemItem.<CoordinateSystem>.<Name>.Value
                NewSystem.CoordinateSystem.Author = systemItem.<CoordinateSystem>.<Author>.Value
                NewSystem.CoordinateSystem.Code = systemItem.<CoordinateSystem>.<Code>.Value
                NewSystem.CoordinateSystem.Type = systemItem.<CoordinateSystem>.<Type>.Value

                List.Add(NewSystem)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            CreationDate = XDoc.<CoordinateReferenceSystemList>.<CreationDate>.Value
            LastEditDate = XDoc.<CoordinateReferenceSystemList>.<LastEditDate>.Value
            Description = XDoc.<CoordinateReferenceSystemList>.<Description>.Value

            Dim Systems = From item In XDoc.<CoordinateReferenceSystemList>.<CoordinateReferenceSystem>

            List.Clear()
            For Each systemItem In Systems
                Dim NewSystem As New CoordinateReferenceSystemSummary
                NewSystem.Name = systemItem.<Name>.Value
                NewSystem.Author = systemItem.<Author>.Value
                If systemItem.<Code>.Value = Nothing Then
                    NewSystem.Code = 0
                Else
                    NewSystem.Code = systemItem.<Code>.Value
                End If
                Select Case systemItem.<Type>.Value
                    Case "Compound" '"compound"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Compound
                        NewSystem.Type = CrsTypes.Compound
                    Case "Engineering" '"engineering"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Engineering
                        NewSystem.Type = CrsTypes.Engineering
                    Case "Geocentric" '"geocentric"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geocentric
                        NewSystem.Type = CrsTypes.Geocentric
                    Case "Geographic2D" '"geographic 2D"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic2D
                        NewSystem.Type = CrsTypes.Geographic2D
                    Case "Geographic3D" '"geographic 3D"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic3D
                        NewSystem.Type = CrsTypes.Geographic3D
                    Case "Projected" '"projected"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Projected
                        NewSystem.Type = CrsTypes.Projected
                    Case "Vertical" '"vertical"
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Vertical
                        NewSystem.Type = CrsTypes.Vertical
                    Case Else
                        'NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Unknown
                        NewSystem.Type = CrsTypes.Unknown
                End Select
                NewSystem.Scope = systemItem.<Scope>.Value
                NewSystem.Comments = systemItem.<Comments>.Value
                NewSystem.Deprecated = systemItem.<Deprecated>.Value

                Dim aliasNames = From item In systemItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewSystem.AddAlias(nameItem)
                Next

                'Read Area of Use information: --------------------------------------------------------------
                NewSystem.Area.Name = systemItem.<AreaOfUse>.<Name>.Value
                NewSystem.Area.Author = systemItem.<AreaOfUse>.<Author>.Value
                NewSystem.Area.Code = systemItem.<AreaOfUse>.<Code>.Value
                NewSystem.Area.Type = ""

                'Read Coordinate System information: --------------------------------------------------------------
                NewSystem.CoordinateSystem.Name = systemItem.<CoordinateSystem>.<Name>.Value
                NewSystem.CoordinateSystem.Author = systemItem.<CoordinateSystem>.<Author>.Value
                NewSystem.CoordinateSystem.Code = systemItem.<CoordinateSystem>.<Code>.Value
                NewSystem.CoordinateSystem.Type = systemItem.<CoordinateSystem>.<Type>.Value

                List.Add(NewSystem)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Coordinate Reference System list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Coordinate Reference System table into ds
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System]", myConnection)
            TableName = "[Coordinate Reference System]"
            da.Fill(ds, TableName)

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)

            'Read the Coordinate System table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate System]", myConnection)
            TableName = "Coordinate System"
            da.Fill(ds, TableName)

            Dim expression As String

            ' variables:
            Dim NParams As Integer 'The number of parameters used to define the projection.
            Dim ParamNo As Integer 'The current parameter number.
            Dim ParamCode As Integer 'The parameter code of the current parameter
            'Dim UomCode As Integer 'The unit of measure code corresponding to the parameter
            'Dim TargetUomCode As Integer 'The standard unit code

            Dim NRows As Integer = ds.Tables("[Coordinate Reference System]").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Coordinate Reference System list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Coordinate Reference Systems from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordRefSystems table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewSystem As New CoordinateReferenceSystemSummary
                NewSystem.Name = ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("COORD_REF_SYS_NAME") '.ToString ???
                NewSystem.Author = "EPSG"
                NewSystem.Code = ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("COORD_REF_SYS_CODE")

                Select Case ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("COORD_REF_SYS_KIND")
                    Case "compound"
                        NewSystem.Type = CrsTypes.Compound
                    Case "engineering"
                        NewSystem.Type = CrsTypes.Engineering
                    Case "geocentric"
                        NewSystem.Type = CrsTypes.Geocentric
                    Case "geographic 2D"
                        NewSystem.Type = CrsTypes.Geographic2D
                    Case "geographic 3D"
                        NewSystem.Type = CrsTypes.Geographic3D
                    Case "projected"
                        NewSystem.Type = CrsTypes.Projected
                    Case "vertical"
                        NewSystem.Type = CrsTypes.Vertical
                    Case Else
                        NewSystem.Type = CrsTypes.Unknown
                End Select

                NewSystem.Scope = ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("CRS_SCOPE")

                If IsDBNull(ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("REMARKS")) Then
                    NewSystem.Comments = ""
                Else
                    NewSystem.Comments = ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("REMARKS")
                End If

                NewSystem.Deprecated = ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("DEPRECATED")

                'Add list of alias names --------------------------------------------------------------------------------------------
                expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewSystem.Code)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewSystem.AddAlias(item.Item("ALIAS").ToString)
                Next
                '---------------------------------------------------------------------------------------------------------------------

                'Add Area of Use details ---------------------------------------------------------------------------------------------
                NewSystem.Area.Code = ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("AREA_OF_USE_CODE")
                expression = "[AREA_CODE] = " & Str(NewSystem.Area.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewSystem.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                    NewSystem.Area.Author = "EPSG"
                    NewSystem.Area.Type = ""
                End If
                '---------------------------------------------------------------------------------------------------------------------

                'Add Coordinate System details ---------------------------------------------------------------------------------------
                'Dim coordinateSystem As New XElement("CoordinateSystem")
                'expression = "[COORD_SYS_CODE] = " & Str(CoordSysCode)
                'Dim coordSysParameters = ds.Tables("Coordinate System").Select(expression)
                'If coordSysParameters.Count > 0 Then
                '    Dim csName As New XElement("Name", coordSysParameters(0).Item("COORD_SYS_NAME").ToString)
                '    coordinateSystem.Add(csName)
                '    Dim csAuthor As New XElement("Author", "EPSG")
                '    coordinateSystem.Add(csAuthor)
                '    Dim csCode As New XElement("Code", Str(CoordSysCode))
                '    coordinateSystem.Add(csCode)
                '    Dim csType As New XElement("Type", coordSysParameters(0).Item("COORD_SYS_TYPE").ToString)
                '    coordinateSystem.Add(csType)
                'End If
                'crs.Add(CoordinateSystem)



                If IsDBNull(ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("COORD_SYS_CODE")) Then
                    'No Coordinate Ssytem Code defined
                    NewSystem.CoordinateSystem.Code = 0
                    NewSystem.CoordinateSystem.Name = ""
                    NewSystem.CoordinateSystem.Author = "EPSG"
                    NewSystem.CoordinateSystem.Type = ""
                Else
                    NewSystem.CoordinateSystem.Code = ds.Tables("[Coordinate Reference System]").Rows(RowNo).Item("COORD_SYS_CODE")
                    expression = "[COORD_SYS_CODE] = " & Str(NewSystem.CoordinateSystem.Code)
                    Dim coordSysParameters = ds.Tables("Coordinate System").Select(expression)
                    If coordSysParameters.Count > 0 Then
                        NewSystem.CoordinateSystem.Name = coordSysParameters(0).Item("COORD_SYS_NAME").ToString
                        NewSystem.CoordinateSystem.Author = "EPSG"
                        NewSystem.CoordinateSystem.Type = coordSysParameters(0).Item("COORD_SYS_TYPE").ToString
                    End If

                End If
                '---------------------------------------------------------------------------------------------------------------------


                List.Add(NewSystem)
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Coordinate Reference Systems as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Coordinate Reference System List File-->
                       <CoordinateReferenceSystemList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <CoordinateReferenceSystem>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Type><%= item.Type %></Type>
                                     <Scope><%= item.Scope %></Scope>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                     <CoordinateSystem>
                                         <Name><%= item.CoordinateSystem.Name %></Name>
                                         <Author><%= item.CoordinateSystem.Author %></Author>
                                         <Code><%= item.CoordinateSystem.Code %></Code>
                                         <Type><%= item.CoordinateSystem.Type %></Type>
                                     </CoordinateSystem>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                 </CoordinateReferenceSystem>
                           %>
                       </CoordinateReferenceSystemList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'CoordRefSystemList ------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class CoordSystemList '------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store a list of Coordinate System parameters.

        Public List As New List(Of CoordinateSystem)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the CoordRefSystem List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<CoordinateSystemList>.<CreationDate>.Value
            LastEditDate = XDoc.<CoordinateSystemList>.<LastEditDate>.Value
            Description = XDoc.<CoordinateSystemList>.<Description>.Value

            Dim Systems = From item In XDoc.<CoordinateSystemList>.<CoordinateSystem>

            List.Clear()
            For Each systemItem In Systems
                Dim NewSystem As New CoordinateSystem
                NewSystem.Name = systemItem.<Name>.Value
                NewSystem.Author = systemItem.<Author>.Value
                If systemItem.<Code>.Value = Nothing Then
                    NewSystem.Code = 0
                Else
                    NewSystem.Code = systemItem.<Code>.Value
                End If
                NewSystem.Selected = systemItem.<Selected>.Value
                Select Case systemItem.<Type>.Value
                    'Case "compound"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Compound
                    'Case "engineering"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Engineering
                    'Case "geocentric"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geocentric
                    'Case "geographic 2D"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic2D
                    'Case "geographic 3D"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic3D
                    'Case "projected"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Projected
                    'Case "vertical"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Vertical
                    'Case Else
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Unknown

                    Case "Cartesian"
                        NewSystem.Type = CoordinateSystem.CSTypes.Cartesian
                    Case "Vertical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Vertical
                    Case "Ellipsoidal"
                        NewSystem.Type = CoordinateSystem.CSTypes.Ellipsoidal
                    Case "Spherical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Spherical
                    Case "Affine"
                        NewSystem.Type = CoordinateSystem.CSTypes.Affine
                    Case "Cylindrical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Cylindrical
                    Case "Linear"
                        NewSystem.Type = CoordinateSystem.CSTypes.Linear
                    Case "Polar"
                        NewSystem.Type = CoordinateSystem.CSTypes.Polar
                    Case Else
                        NewSystem.Type = CoordinateSystem.CSTypes.Unknown
                End Select
                NewSystem.Dimension = systemItem.<Dimension>.Value
                NewSystem.Comments = systemItem.<Comments>.Value
                NewSystem.Deprecated = systemItem.<Deprecated>.Value

                Dim aliasNames = From item In systemItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewSystem.AddAlias(nameItem)
                Next

                Dim axisItems = From item In systemItem.<AxisList>.<Axis>

                For Each axisItem In axisItems
                    Dim NewAxis As New CoordinateAxis
                    NewAxis.Name = axisItem.<Name>.Value
                    NewAxis.Author = axisItem.<Author>.Value
                    NewAxis.Code = axisItem.<Code>.Value
                    NewAxis.Description = axisItem.<Description>.Value
                    NewAxis.Comments = axisItem.<Comments>.Value
                    NewAxis.Orientation = axisItem.<Orientation>.Value
                    NewAxis.Order = axisItem.<Order>.Value
                    If axisItem.<Deprecated>.Value = "" Then
                        NewAxis.Deprecated = False
                    Else
                        NewAxis.Deprecated = axisItem.<Deprecated>.Value
                    End If
                    'NewAxis.Deprecated = axisItem.<Deprecated>.Value
                    NewAxis.UnitOfMeasure.Name = axisItem.<Unit>.<Name>.Value
                    NewAxis.UnitOfMeasure.Code = axisItem.<Unit>.<Code>.Value
                    'NewAxis.UnitOfMeasure.Type = axisItem.<Unit>.<Type>.Value
                    Select Case axisItem.<Unit>.<Type>.Value
                        Case "angle"
                            'NewAxis.UnitOfMeasure.Type = TDS_Utilities.Coordinates.clsUnitOfMeasure.EnumUOMType.Angle
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Angle
                        Case "length"
                            'NewAxis.UnitOfMeasure.Type = TDS_Utilities.Coordinates.clsUnitOfMeasure.EnumUOMType.Length
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Length
                        Case "scale"
                            'NewAxis.UnitOfMeasure.Type = TDS_Utilities.Coordinates.clsUnitOfMeasure.EnumUOMType.Scale
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Scale
                        Case "time"
                            'NewAxis.UnitOfMeasure.Type = TDS_Utilities.Coordinates.clsUnitOfMeasure.EnumUOMType.Time
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Time
                        Case Else
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Unknown
                    End Select
                    NewAxis.UnitOfMeasure.StandardUnitName = axisItem.<Unit>.<TargetUnit>.Value
                    NewAxis.UnitOfMeasure.FactorB = axisItem.<Unit>.<FactorB>.Value
                    NewAxis.UnitOfMeasure.FactorC = axisItem.<Unit>.<FactorC>.Value
                    NewAxis.UnitOfMeasure.Comments = axisItem.<Unit>.<Comments>.Value
                    NewAxis.UnitOfMeasure.Deprecated = axisItem.<Unit>.<Deprecated>.Value
                    NewSystem.Axis.Add(NewAxis)
                Next

                List.Add(NewSystem)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<CoordinateSystemList>.<CreationDate>.Value
            LastEditDate = XDoc.<CoordinateSystemList>.<LastEditDate>.Value
            Description = XDoc.<CoordinateSystemList>.<Description>.Value

            Dim Systems = From item In XDoc.<CoordinateSystemList>.<CoordinateSystem>

            List.Clear()
            For Each systemItem In Systems
                Dim NewSystem As New CoordinateSystem
                NewSystem.Name = systemItem.<Name>.Value
                NewSystem.Author = systemItem.<Author>.Value
                If systemItem.<Code>.Value = Nothing Then
                    NewSystem.Code = 0
                Else
                    NewSystem.Code = systemItem.<Code>.Value
                End If
                NewSystem.Selected = systemItem.<Selected>.Value
                Select Case systemItem.<Type>.Value
                    'Case "compound"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Compound
                    'Case "engineering"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Engineering
                    'Case "geocentric"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geocentric
                    'Case "geographic 2D"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic2D
                    'Case "geographic 3D"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic3D
                    'Case "projected"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Projected
                    'Case "vertical"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Vertical
                    'Case Else
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Unknown

                    Case "Cartesian"
                        NewSystem.Type = CoordinateSystem.CSTypes.Cartesian
                    Case "Vertical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Vertical
                    Case "Ellipsoidal"
                        NewSystem.Type = CoordinateSystem.CSTypes.Ellipsoidal
                    Case "Spherical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Spherical
                    Case "Affine"
                        NewSystem.Type = CoordinateSystem.CSTypes.Affine
                    Case "Cylindrical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Cylindrical
                    Case "Linear"
                        NewSystem.Type = CoordinateSystem.CSTypes.Linear
                    Case "Polar"
                        NewSystem.Type = CoordinateSystem.CSTypes.Polar
                    Case Else
                        NewSystem.Type = CoordinateSystem.CSTypes.Unknown
                End Select
                NewSystem.Dimension = systemItem.<Dimension>.Value
                NewSystem.Comments = systemItem.<Comments>.Value
                NewSystem.Deprecated = systemItem.<Deprecated>.Value

                Dim aliasNames = From item In systemItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewSystem.AddAlias(nameItem)
                Next

                Dim axisItems = From item In systemItem.<AxisList>.<Axis>

                For Each axisItem In axisItems
                    Dim NewAxis As New CoordinateAxis
                    NewAxis.Name = axisItem.<Name>.Value
                    NewAxis.Author = axisItem.<Author>.Value
                    NewAxis.Code = axisItem.<Code>.Value
                    NewAxis.Description = axisItem.<Description>.Value
                    NewAxis.Comments = axisItem.<Comments>.Value
                    NewAxis.Orientation = axisItem.<Orientation>.Value
                    NewAxis.Order = axisItem.<Order>.Value
                    If axisItem.<Deprecated>.Value = "" Then
                        NewAxis.Deprecated = False
                    Else
                        NewAxis.Deprecated = axisItem.<Deprecated>.Value
                    End If
                    'NewAxis.Deprecated = axisItem.<Deprecated>.Value
                    NewAxis.UnitOfMeasure.Name = axisItem.<Unit>.<Name>.Value
                    NewAxis.UnitOfMeasure.Code = axisItem.<Unit>.<Code>.Value
                    'NewAxis.UnitOfMeasure.Type = axisItem.<Unit>.<Type>.Value
                    Select Case axisItem.<Unit>.<Type>.Value
                        Case "angle"
                            'NewAxis.UnitOfMeasure.Type = TDS_Utilities.Coordinates.clsUnitOfMeasure.EnumUOMType.Angle
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Angle
                        Case "length"
                            'NewAxis.UnitOfMeasure.Type = TDS_Utilities.Coordinates.clsUnitOfMeasure.EnumUOMType.Length
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Length
                        Case "scale"
                            'NewAxis.UnitOfMeasure.Type = TDS_Utilities.Coordinates.clsUnitOfMeasure.EnumUOMType.Scale
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Scale
                        Case "time"
                            'NewAxis.UnitOfMeasure.Type = TDS_Utilities.Coordinates.clsUnitOfMeasure.EnumUOMType.Time
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Time
                        Case Else
                            NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Unknown
                    End Select
                    NewAxis.UnitOfMeasure.StandardUnitName = axisItem.<Unit>.<TargetUnit>.Value
                    NewAxis.UnitOfMeasure.FactorB = axisItem.<Unit>.<FactorB>.Value
                    NewAxis.UnitOfMeasure.FactorC = axisItem.<Unit>.<FactorC>.Value
                    NewAxis.UnitOfMeasure.Comments = axisItem.<Unit>.<Comments>.Value
                    NewAxis.UnitOfMeasure.Deprecated = axisItem.<Unit>.<Deprecated>.Value
                    NewSystem.Axis.Add(NewAxis)
                Next

                List.Add(NewSystem)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Coordinate System list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Coordinate System table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate System]", myConnection)
            TableName = "CoordSys"
            da.Fill(ds, TableName)
            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Coordinate Axis table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate Axis]", myConnection)
            TableName = "Axis"
            da.Fill(ds, TableName)
            'COORD_AXIS_CODE COORD_SYS_CODE COORD_AXIS_NAME_CODE COORD_AXIS_ORIENTATION COORD_AXIS_ABBREVIATION UOM_CODE ORDER

            'Read the Coordinate Axis Name table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate Axis Name]", myConnection)
            TableName = "AxisName"
            da.Fill(ds, TableName)
            'COORD_AXIS_NAME_CODE COORD_AXIS_NAME DESCRIPTION REMARKS DEPRECATED

            'Read the Unit of Measure table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            TableName = "Unit"
            da.Fill(ds, TableName)
            'UOM_CODE UNIT_OF_MEAS_NAME UNIT_OF_MEAS_TYPE TARGET_UOM_CODE FACTOR_B FACTOR_C REMARKS DEPRECATED

            Dim expression As String






            ' variables:
            Dim NParams As Integer 'The number of parameters used to define the projection.
            Dim ParamNo As Integer 'The current parameter number.
            Dim ParamCode As Integer 'The parameter code of the current parameter
            'Dim UomCode As Integer 'The unit of measure code corresponding to the parameter
            Dim TargetUomCode As Integer 'The standard unit code

            Dim NRows As Integer = ds.Tables("CoordSys").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Coordinate System list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Coordinate Systems from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewSystem As New CoordinateSystem
                NewSystem.Name = ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME")
                NewSystem.Author = "EPSG"
                NewSystem.Code = ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_CODE")

                'NOTE: Coordinate System types are: Cartesian, vertical, ellipsoidal, spherical
                Select Case ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_TYPE")
                    'Case "compound"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Compound
                    'Case "engineering"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Engineering
                    'Case "geocentric"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geocentric
                    'Case "geographic 2D"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic2D
                    'Case "geographic 3D"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Geographic3D
                    'Case "projected"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Projected
                    'Case "vertical"
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Vertical
                    'Case Else
                    '    NewSystem.Type = CoordinateReferenceSystemSummary.CrsTypes.Unknown

                    Case "Cartesian"
                        NewSystem.Type = CoordinateSystem.CSTypes.Cartesian
                    Case "vertical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Vertical
                    Case "ellipsoidal"
                        NewSystem.Type = CoordinateSystem.CSTypes.Ellipsoidal
                    Case "spherical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Spherical
                    Case "affine"
                        NewSystem.Type = CoordinateSystem.CSTypes.Affine
                    Case "cylindrical"
                        NewSystem.Type = CoordinateSystem.CSTypes.Cylindrical
                    Case "linear"
                        NewSystem.Type = CoordinateSystem.CSTypes.Linear
                    Case "polar"
                        NewSystem.Type = CoordinateSystem.CSTypes.Polar
                    Case Else
                        NewSystem.Type = CoordinateSystem.CSTypes.Unknown
                End Select

                NewSystem.Dimension = ds.Tables("CoordSys").Rows(RowNo).Item("DIMENSION")

                If IsDBNull(ds.Tables("CoordSys").Rows(RowNo).Item("REMARKS")) Then
                    NewSystem.Comments = ""
                Else
                    NewSystem.Comments = ds.Tables("CoordSys").Rows(RowNo).Item("REMARKS")
                End If

                NewSystem.Deprecated = ds.Tables("CoordSys").Rows(RowNo).Item("DEPRECATED")

                'Add list of alias names --------------------------------------------------------------------------------------------
                expression = "[OBJECT_TABLE_NAME] = 'Coordinate System' AND [OBJECT_CODE] = " & Str(NewSystem.Code)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewSystem.AddAlias(item.Item("ALIAS").ToString)
                Next
                '---------------------------------------------------------------------------------------------------------------------

                'Add Axis list ---------------------------------------------------------------------------------------------
                expression = "[COORD_SYS_CODE] = " & Str(NewSystem.Code)
                Dim axisResult = ds.Tables("Axis").Select(expression)
                For Each item In axisResult
                    Dim NewAxis As New CoordinateAxis
                    NewAxis.Code = item.Item("COORD_AXIS_NAME_CODE")
                    expression = "[COORD_AXIS_NAME_CODE] = " & Str(NewAxis.Code)
                    Dim axisNameResult = ds.Tables("AxisName").Select(expression)
                    If axisNameResult.Count = 0 Then
                        RaiseEvent ErrorMessage("Missing axis name. Coordinate System name: " & ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME") & vbCrLf)
                        NewAxis.Name = ""
                        NewAxis.Description = ""
                        NewAxis.Comments = ""
                    ElseIf axisNameResult.Count > 1 Then
                        RaiseEvent ErrorMessage("Greater than one axis name. Coordinate System name: " & ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME") & vbCrLf)
                        NewAxis.Name = axisNameResult(0).Item("COORD_AXIS_NAME").ToString
                        NewAxis.Description = axisNameResult(0).Item("DESCRIPTION").ToString
                        NewAxis.Comments = axisNameResult(0).Item("REMARKS").ToString
                    Else
                        NewAxis.Name = axisNameResult(0).Item("COORD_AXIS_NAME").ToString
                        NewAxis.Description = axisNameResult(0).Item("DESCRIPTION").ToString
                        NewAxis.Comments = axisNameResult(0).Item("REMARKS").ToString
                    End If
                    NewAxis.Orientation = item.Item("COORD_AXIS_ORIENTATION").ToString
                    NewAxis.Abbreviation = item.Item("COORD_AXIS_ABBREVIATION").ToString
                    NewAxis.Order = item.Item("ORDER").ToString

                    NewAxis.UnitOfMeasure.Code = item.Item("UOM_CODE")
                    expression = "[UOM_CODE] = " & Str(NewAxis.UnitOfMeasure.Code)
                    Dim axisUnitResult = ds.Tables("Unit").Select(expression)
                    If axisUnitResult.Count = 0 Then
                        RaiseEvent ErrorMessage("Missing axis unit. Coordinate System name: " & ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME") & vbCrLf)
                    ElseIf axisUnitResult.Count > 1 Then
                        RaiseEvent ErrorMessage("Greater than one axis unit. Coordinate System name: " & ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME") & vbCrLf)
                    Else
                        NewAxis.UnitOfMeasure.Name = axisUnitResult(0).Item("UNIT_OF_MEAS_NAME").ToString
                        Select Case axisUnitResult(0).Item("UNIT_OF_MEAS_TYPE").ToString
                            Case "Angle"
                                NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Angle
                            Case "Length"
                                NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Length
                            Case "Scale"
                                NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Scale
                            Case "Time"
                                NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Time
                            Case Else
                                NewAxis.UnitOfMeasure.Type = UnitOfMeasure.UOMTypes.Unknown
                        End Select

                        TargetUomCode = axisUnitResult(0).Item("TARGET_UOM_CODE")
                        expression = "[UOM_CODE] = " & Str(TargetUomCode)
                        Dim targetUnitResult = ds.Tables("Unit").Select(expression)
                        If targetUnitResult.Count = 0 Then
                            RaiseEvent ErrorMessage("Missing axis target unit. Coordinate System name: " & ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME") & vbCrLf)
                        ElseIf targetUnitResult.Count > 1 Then
                            RaiseEvent ErrorMessage("Greater than one axis target unit. Coordinate System name: " & ds.Tables("CoordSys").Rows(RowNo).Item("COORD_SYS_NAME") & vbCrLf)
                        Else
                            NewAxis.UnitOfMeasure.StandardUnitName = targetUnitResult(0).Item("UNIT_OF_MEAS_NAME").ToString
                            If IsDBNull(axisUnitResult(0).Item("FACTOR_B")) Then
                                NewAxis.UnitOfMeasure.FactorB = 1
                            Else
                                NewAxis.UnitOfMeasure.FactorB = axisUnitResult(0).Item("FACTOR_B")
                            End If
                            If IsDBNull(axisUnitResult(0).Item("FACTOR_C")) Then
                                NewAxis.UnitOfMeasure.FactorC = 1
                            Else
                                NewAxis.UnitOfMeasure.FactorC = axisUnitResult(0).Item("FACTOR_C")
                            End If


                            If IsDBNull(axisUnitResult(0).Item("REMARKS")) Then
                                NewAxis.UnitOfMeasure.Comments = ""
                            Else
                                NewAxis.UnitOfMeasure.Comments = axisUnitResult(0).Item("REMARKS")
                            End If

                            NewAxis.UnitOfMeasure.Deprecated = axisUnitResult(0).Item("DEPRECATED")
                        End If
                    End If
                    NewSystem.Axis.Add(NewAxis)
                Next
                List.Add(NewSystem)
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Coordinate Systems as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Coordinate System List File-->
                       <CoordinateSystemList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <CoordinateSystem>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Type><%= item.Type %></Type>
                                     <Dimension><%= item.Dimension %></Dimension>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <AxisList>
                                         <%= From axisItem In item.Axis _
                                             Select _
                                             <Axis>
                                                 <Name><%= axisItem.Name %></Name>
                                                 <Description><%= axisItem.Description %></Description>
                                                 <Comments><%= axisItem.Comments %></Comments>
                                                 <Orientation><%= axisItem.Orientation %></Orientation>
                                                 <Abbreviation><%= axisItem.Abbreviation %></Abbreviation>
                                                 <Order><%= axisItem.Order %></Order>
                                                 <Unit>
                                                     <Name><%= axisItem.UnitOfMeasure.Name %></Name>
                                                     <Code><%= axisItem.UnitOfMeasure.Code %></Code>
                                                     <Type><%= axisItem.UnitOfMeasure.Type %></Type>
                                                     <StandardUnitName><%= axisItem.UnitOfMeasure.StandardUnitName %></StandardUnitName>
                                                     <FactorB><%= axisItem.UnitOfMeasure.FactorB %></FactorB>
                                                     <FactorC><%= axisItem.UnitOfMeasure.FactorC %></FactorC>
                                                     <Comments><%= axisItem.UnitOfMeasure.Comments %></Comments>
                                                     <Deprecated><%= axisItem.UnitOfMeasure.Deprecated %></Deprecated>
                                                 </Unit>
                                             </Axis>
                                         %>
                                     </AxisList>
                                 </CoordinateSystem>
                           %>
                           <!---->
                       </CoordinateSystemList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'CoordSystemList ---------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class DatumList '------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store a list of Datum parameters.

        Public List As New List(Of DatumSummaryWithArea)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the CoordRefSystem List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<DatumList>.<CreationDate>.Value
            LastEditDate = XDoc.<DatumList>.<LastEditDate>.Value
            Description = XDoc.<DatumList>.<Description>.Value

            Dim Datums = From item In XDoc.<DatumList>.<Datum>

            List.Clear()
            For Each datumItem In Datums
                Dim NewDatum As New DatumSummaryWithArea
                NewDatum.Name = datumItem.<Name>.Value
                NewDatum.Author = datumItem.<Author>.Value
                If datumItem.<Code>.Value = Nothing Then
                    NewDatum.Code = 0
                Else
                    NewDatum.Code = datumItem.<Code>.Value
                End If
                NewDatum.Selected = datumItem.<Selected>.Value

                If datumItem.<Type>.Value = Nothing Then
                    NewDatum.Type = DatumSummary.DatumTypes.Unknown 'TDS_Utilities.Coordinates.clsDatumSummary.EnumDatumType.Unknown
                Else
                    Select Case datumItem.<Type>.Value
                        Case "geodetic"
                            NewDatum.Type = DatumSummary.DatumTypes.Geodetic
                        Case "vertical"
                            NewDatum.Type = DatumSummary.DatumTypes.Vertical
                        Case "engineering"
                            NewDatum.Type = DatumSummary.DatumTypes.Engineering
                        Case "image"
                            NewDatum.Type = DatumSummary.DatumTypes.Image
                        Case Else
                            NewDatum.Type = DatumSummary.DatumTypes.Unknown
                    End Select
                End If

                NewDatum.OriginDescription = datumItem.<Origin>.Value
                NewDatum.Epoch = datumItem.<Epoch>.Value
                NewDatum.Scope = datumItem.<Scope>.Value
                NewDatum.Comments = datumItem.<Comments>.Value
                NewDatum.Deprecated = datumItem.<Deprecated>.Value

                Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewDatum.AddAlias(nameItem)
                Next

                'Read area of use information: -------------------------------------------------------------------------------------------------
                NewDatum.AreaOfUse.Name = datumItem.<AreaOfUse>.<Name>.Value
                NewDatum.AreaOfUse.Code = datumItem.<AreaOfUse>.<Code>.Value
                NewDatum.AreaOfUse.Comments = datumItem.<AreaOfUse>.<Comments>.Value
                NewDatum.AreaOfUse.Description = datumItem.<AreaOfUse>.<Description>.Value
                If datumItem.<AreaOfUse>.<SouthLatitude>.Value = "" Then
                    NewDatum.AreaOfUse.SouthLatitude = 0
                Else
                    NewDatum.AreaOfUse.SouthLatitude = datumItem.<AreaOfUse>.<SouthLatitude>.Value
                End If
                If datumItem.<AreaOfUse>.<NorthLatitude>.Value = "" Then
                    NewDatum.AreaOfUse.NorthLatitude = 0
                Else
                    NewDatum.AreaOfUse.NorthLatitude = datumItem.<AreaOfUse>.<NorthLatitude>.Value
                End If
                If datumItem.<AreaOfUse>.<WestLongitude>.Value = "" Then
                    NewDatum.AreaOfUse.WestLongitude = 0
                Else
                    NewDatum.AreaOfUse.WestLongitude = datumItem.<AreaOfUse>.<WestLongitude>.Value
                End If
                If datumItem.<AreaOfUse>.<EastLongitude>.Value = "" Then
                    NewDatum.AreaOfUse.EastLongitude = 0
                Else
                    NewDatum.AreaOfUse.EastLongitude = datumItem.<AreaOfUse>.<EastLongitude>.Value
                End If

                NewDatum.AreaOfUse.IsoA2Code = datumItem.<AreaOfUse>.<IsoA2Code>.Value
                NewDatum.AreaOfUse.IsoA3Code = datumItem.<AreaOfUse>.<IsoA3Code>.Value
                If datumItem.<AreaOfUse>.<IsoNCode>.Value = "" Then
                    NewDatum.AreaOfUse.IsoNCode = 0
                Else
                    NewDatum.AreaOfUse.IsoNCode = datumItem.<AreaOfUse>.<IsoNCode>.Value
                End If

                Dim aliasAOUNames = From item In datumItem.<AreaOfUse>.<AliasNames>.<AliasName>
                For Each nameItem In aliasAOUNames
                    NewDatum.AreaOfUse.AddAlias(nameItem)
                Next

                List.Add(NewDatum)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<DatumList>.<CreationDate>.Value
            LastEditDate = XDoc.<DatumList>.<LastEditDate>.Value
            Description = XDoc.<DatumList>.<Description>.Value

            Dim Datums = From item In XDoc.<DatumList>.<Datum>

            List.Clear()
            For Each datumItem In Datums
                Dim NewDatum As New DatumSummaryWithArea
                NewDatum.Name = datumItem.<Name>.Value
                NewDatum.Author = datumItem.<Author>.Value
                If datumItem.<Code>.Value = Nothing Then
                    NewDatum.Code = 0
                Else
                    NewDatum.Code = datumItem.<Code>.Value
                End If
                NewDatum.Selected = datumItem.<Selected>.Value

                If datumItem.<Type>.Value = Nothing Then
                    NewDatum.Type = DatumSummary.DatumTypes.Unknown 'TDS_Utilities.Coordinates.clsDatumSummary.EnumDatumType.Unknown
                Else
                    Select Case datumItem.<Type>.Value
                        Case "geodetic"
                            NewDatum.Type = DatumSummary.DatumTypes.Geodetic
                        Case "vertical"
                            NewDatum.Type = DatumSummary.DatumTypes.Vertical
                        Case "engineering"
                            NewDatum.Type = DatumSummary.DatumTypes.Engineering
                        Case "image"
                            NewDatum.Type = DatumSummary.DatumTypes.Image
                        Case Else
                            NewDatum.Type = DatumSummary.DatumTypes.Unknown
                    End Select
                End If

                NewDatum.OriginDescription = datumItem.<Origin>.Value
                NewDatum.Epoch = datumItem.<Epoch>.Value
                NewDatum.Scope = datumItem.<Scope>.Value
                NewDatum.Comments = datumItem.<Comments>.Value
                NewDatum.Deprecated = datumItem.<Deprecated>.Value

                Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewDatum.AddAlias(nameItem)
                Next

                'Read area of use information: -------------------------------------------------------------------------------------------------
                NewDatum.AreaOfUse.Name = datumItem.<AreaOfUse>.<Name>.Value
                NewDatum.AreaOfUse.Code = datumItem.<AreaOfUse>.<Code>.Value
                NewDatum.AreaOfUse.Comments = datumItem.<AreaOfUse>.<Comments>.Value
                NewDatum.AreaOfUse.Description = datumItem.<AreaOfUse>.<Description>.Value
                If datumItem.<AreaOfUse>.<SouthLatitude>.Value = "" Then
                    NewDatum.AreaOfUse.SouthLatitude = 0
                Else
                    NewDatum.AreaOfUse.SouthLatitude = datumItem.<AreaOfUse>.<SouthLatitude>.Value
                End If
                If datumItem.<AreaOfUse>.<NorthLatitude>.Value = "" Then
                    NewDatum.AreaOfUse.NorthLatitude = 0
                Else
                    NewDatum.AreaOfUse.NorthLatitude = datumItem.<AreaOfUse>.<NorthLatitude>.Value
                End If
                If datumItem.<AreaOfUse>.<WestLongitude>.Value = "" Then
                    NewDatum.AreaOfUse.WestLongitude = 0
                Else
                    NewDatum.AreaOfUse.WestLongitude = datumItem.<AreaOfUse>.<WestLongitude>.Value
                End If
                If datumItem.<AreaOfUse>.<EastLongitude>.Value = "" Then
                    NewDatum.AreaOfUse.EastLongitude = 0
                Else
                    NewDatum.AreaOfUse.EastLongitude = datumItem.<AreaOfUse>.<EastLongitude>.Value
                End If

                NewDatum.AreaOfUse.IsoA2Code = datumItem.<AreaOfUse>.<IsoA2Code>.Value
                NewDatum.AreaOfUse.IsoA3Code = datumItem.<AreaOfUse>.<IsoA3Code>.Value
                If datumItem.<AreaOfUse>.<IsoNCode>.Value = "" Then
                    NewDatum.AreaOfUse.IsoNCode = 0
                Else
                    NewDatum.AreaOfUse.IsoNCode = datumItem.<AreaOfUse>.<IsoNCode>.Value
                End If

                Dim aliasAOUNames = From item In datumItem.<AreaOfUse>.<AliasNames>.<AliasName>
                For Each nameItem In aliasAOUNames
                    NewDatum.AreaOfUse.AddAlias(nameItem)
                Next

                List.Add(NewDatum)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Datum list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Datum table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Datum]", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)


            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)

            Dim expression As String

            ' variables:
            Dim NParams As Integer 'The number of parameters used to define the projection.
            Dim ParamNo As Integer 'The current parameter number.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number


            Dim NRows As Integer = ds.Tables("Datum").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Datum list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Datums from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewDatum As New DatumSummaryWithArea
                NewDatum.Name = ds.Tables("Datum").Rows(RowNo).Item("DATUM_NAME")
                NewDatum.Author = "EPSG"
                NewDatum.Code = ds.Tables("Datum").Rows(RowNo).Item("DATUM_CODE")
                NewDatum.AreaOfUse.Code = ds.Tables("Datum").Rows(RowNo).Item("AREA_OF_USE_CODE")

                Select Case ds.Tables("Datum").Rows(RowNo).Item("DATUM_TYPE")
                    Case "geodetic"
                        NewDatum.Type = DatumSummary.DatumTypes.Geodetic
                    Case "vertical"
                        NewDatum.Type = DatumSummary.DatumTypes.Vertical
                    Case "engineering"
                        NewDatum.Type = DatumSummary.DatumTypes.Engineering
                    Case "image"
                        NewDatum.Type = DatumSummary.DatumTypes.Image
                    Case Else
                        NewDatum.Type = DatumSummary.DatumTypes.Unknown
                End Select

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")) Then
                    NewDatum.OriginDescription = ""
                Else
                    NewDatum.OriginDescription = ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")
                End If

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")) Then
                    NewDatum.Epoch = ""
                Else
                    NewDatum.Epoch = ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")
                End If

                NewDatum.Scope = ds.Tables("Datum").Rows(RowNo).Item("DATUM_SCOPE")

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REMARKS")) Then
                    NewDatum.Comments = ""
                Else
                    NewDatum.Comments = ds.Tables("Datum").Rows(RowNo).Item("REMARKS")
                End If

                NewDatum.Deprecated = ds.Tables("Datum").Rows(RowNo).Item("DEPRECATED")

                'Add list of alias names --------------------------------------------------------------------------------------------
                expression = "[OBJECT_TABLE_NAME] = 'Datum' AND [OBJECT_CODE] = " & Str(NewDatum.Code)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewDatum.AddAlias(item.Item("ALIAS").ToString)
                Next
                '---------------------------------------------------------------------------------------------------------------------

                'Add Area of Use details ---------------------------------------------------------------------------------------------
                expression = "[AREA_CODE] = " & Str(NewDatum.AreaOfUse.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewDatum.AreaOfUse.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                    NewDatum.AreaOfUse.Author = "EPSG"
                    NewDatum.AreaOfUse.Description = areaOfUseParameters(0).Item("AREA_OF_USE").ToString
                    NewDatum.AreaOfUse.Comments = areaOfUseParameters(0).Item("REMARKS").ToString
                    If IsDBNull(areaOfUseParameters(0).Item("AREA_SOUTH_BOUND_LAT")) Then
                        NewDatum.AreaOfUse.SouthLatitude = Double.NaN
                    Else
                        NewDatum.AreaOfUse.SouthLatitude = areaOfUseParameters(0).Item("AREA_SOUTH_BOUND_LAT")
                    End If
                    If IsDBNull(areaOfUseParameters(0).Item("AREA_NORTH_BOUND_LAT")) Then
                        NewDatum.AreaOfUse.NorthLatitude = Double.NaN
                    Else
                        NewDatum.AreaOfUse.NorthLatitude = areaOfUseParameters(0).Item("AREA_NORTH_BOUND_LAT")
                    End If
                    If IsDBNull(areaOfUseParameters(0).Item("AREA_WEST_BOUND_LON")) Then
                        NewDatum.AreaOfUse.WestLongitude = Double.NaN
                    Else
                        NewDatum.AreaOfUse.WestLongitude = areaOfUseParameters(0).Item("AREA_WEST_BOUND_LON")
                    End If
                    If IsDBNull(areaOfUseParameters(0).Item("AREA_EAST_BOUND_LON")) Then
                        NewDatum.AreaOfUse.EastLongitude = Double.NaN
                    Else
                        NewDatum.AreaOfUse.EastLongitude = areaOfUseParameters(0).Item("AREA_EAST_BOUND_LON")
                    End If

                    If IsDBNull(areaOfUseParameters(0).Item("ISO_A2_CODE")) Then
                        NewDatum.AreaOfUse.IsoA2Code = 0
                    Else
                        NewDatum.AreaOfUse.IsoA2Code = areaOfUseParameters(0).Item("ISO_A2_CODE")
                    End If
                    If IsDBNull(areaOfUseParameters(0).Item("ISO_A3_CODE")) Then
                        NewDatum.AreaOfUse.IsoA3Code = 0
                    Else
                        NewDatum.AreaOfUse.IsoA3Code = areaOfUseParameters(0).Item("ISO_A3_CODE")
                    End If
                    If IsDBNull(areaOfUseParameters(0).Item("ISO_N_CODE")) Then
                        NewDatum.AreaOfUse.IsoNCode = 0
                    Else
                        NewDatum.AreaOfUse.IsoNCode = areaOfUseParameters(0).Item("ISO_N_CODE")
                    End If

                    'Add list of alias names --------------------------------------------------------------------------------------------
                    Dim aouAliasNames As New XElement("AliasNames")
                    'expression = "[OBJECT_TABLE_NAME] = 'Area' AND [OBJECT_CODE] = " & areaOfUseParameters(0).Item("AREA_CODE").ToString
                    expression = "[OBJECT_TABLE_NAME] = 'Area' AND [OBJECT_CODE] = " & Str(NewDatum.AreaOfUse.Code)
                    Dim result2 = ds.Tables("Alias").Select(expression)
                    For Each item In result2
                        NewDatum.AreaOfUse.AliasName.Add(item.Item("ALIAS").ToString)
                    Next
                End If
                List.Add(NewDatum)
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Coordinate Systems as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Coordinate System List File-->
                       <DatumList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <Datum>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Type><%= item.Type %></Type>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <OriginDescription><%= item.OriginDescription %></OriginDescription>
                                     <Epoch><%= item.Epoch %></Epoch>
                                     <Scope><%= item.Scope %></Scope>
                                     <AreaOfUse>
                                         <Name><%= item.AreaOfUse.Name %></Name>
                                         <Code><%= item.AreaOfUse.Code %></Code>
                                         <Author><%= item.AreaOfUse.Author %></Author>
                                         <Description><%= item.AreaOfUse.Description %></Description>
                                         <Comments><%= item.AreaOfUse.Comments %></Comments>
                                         <SouthLatitude><%= item.AreaOfUse.SouthLatitude %></SouthLatitude>
                                         <NorthLatitude><%= item.AreaOfUse.NorthLatitude %></NorthLatitude>
                                         <WestLongitude><%= item.AreaOfUse.WestLongitude %></WestLongitude>
                                         <EastLongitude><%= item.AreaOfUse.EastLongitude %></EastLongitude>
                                         <IsoA2Code><%= item.AreaOfUse.IsoA2Code %></IsoA2Code>
                                         <IsoA3Code><%= item.AreaOfUse.IsoA3Code %></IsoA3Code>
                                         <IsoNCode><%= item.AreaOfUse.IsoNCode %></IsoNCode>
                                         <AliasNames>
                                             <%= From areaNameItem In item.AreaOfUse.AliasName _
                                                 Select _
                                                 <AliasName><%= areaNameItem %></AliasName>
                                             %>
                                         </AliasNames>
                                     </AreaOfUse>
                                 </Datum>
                           %>
                           <!---->
                       </DatumList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'DatumList ---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class TransformationList '---------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store a list of Transformation parameters.

        Public List As New List(Of Transformation)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Transformation List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<TransformationList>.<CreationDate>.Value
            LastEditDate = XDoc.<TransformationList>.<LastEditDate>.Value
            Description = XDoc.<TransformationList>.<Description>.Value

            Dim Transformations = From item In XDoc.<TransformationList>.<Transformation>

            List.Clear()
            For Each transItem In Transformations
                Dim NewTrans As New Transformation
                NewTrans.Name = transItem.<Name>.Value
                NewTrans.Author = transItem.<Author>.Value
                If transItem.<Code>.Value = Nothing Then
                    NewTrans.Code = 0
                Else
                    NewTrans.Code = transItem.<Code>.Value
                End If
                NewTrans.Selected = transItem.<Selected>.Value

                NewTrans.Version = transItem.<Version>.Value
                NewTrans.VariantNo = transItem.<Variant>.Value
                NewTrans.Scope = transItem.<Scope>.Value
                NewTrans.Accuracy = transItem.<Accuracy>.Value
                NewTrans.Comments = transItem.<Comments>.Value
                NewTrans.Deprecated = transItem.<Deprecated>.Value

                Dim aliasNames = From item In transItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewTrans.AddAlias(nameItem)
                Next

                'Read Source CRS information: -------------------------------------------------------------------------------------------------
                NewTrans.SourceCRS.Name = transItem.<SourceCRS>.<Name>.Value
                NewTrans.SourceCRS.Author = transItem.<SourceCRS>.<Author>.Value
                NewTrans.SourceCRS.Code = transItem.<SourceCRS>.<Code>.Value

                'Read Target CRS information: -------------------------------------------------------------------------------------------------
                NewTrans.TargetCRS.Name = transItem.<TargetCRS>.<Name>.Value
                NewTrans.TargetCRS.Author = transItem.<TargetCRS>.<Author>.Value
                NewTrans.TargetCRS.Code = transItem.<TargetCRS>.<Code>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewTrans.Area.Name = transItem.<AreaOfUse>.<Name>.Value
                NewTrans.Area.Author = transItem.<AreaOfUse>.<Author>.Value
                NewTrans.Area.Code = transItem.<AreaOfUse>.<Code>.Value

                NewTrans.ReverseOp = transItem.<ReverseOp>.Value

                'Read Method information: -------------------------------------------------------------------------------------------------
                NewTrans.Method.Name = transItem.<Method>.<Name>.Value
                NewTrans.Method.Author = transItem.<Method>.<Author>.Value
                NewTrans.Method.Code = transItem.<Method>.<Code>.Value

                'Read Source Coord Diff Unit information: -------------------------------------------------------------------------------------------------
                NewTrans.SourceCoordDiffUnit.Name = transItem.<SourceCoordDiffUnit>.<Name>.Value
                NewTrans.SourceCoordDiffUnit.Author = transItem.<SourceCoordDiffUnit>.<Author>.Value
                NewTrans.SourceCoordDiffUnit.Code = transItem.<SourceCoordDiffUnit>.<Code>.Value

                'Read Target Coord Diff Unit information: -------------------------------------------------------------------------------------------------
                NewTrans.TargetCoordDiffUnit.Name = transItem.<TargetCoordDiffUnit>.<Name>.Value
                NewTrans.TargetCoordDiffUnit.Author = transItem.<TargetCoordDiffUnit>.<Author>.Value
                NewTrans.TargetCoordDiffUnit.Code = transItem.<TargetCoordDiffUnit>.<Code>.Value

                'Read Parameter List
                Dim parameters = From item In transItem.<ParameterList>.<Parameter>

                For Each parameterItem In parameters
                    Dim NewParameter As New ValueSummary
                    NewParameter.Name = parameterItem.<Name>.Value
                    NewParameter.Value = parameterItem.<Value>.Value
                    NewParameter.Unit.Name = parameterItem.<UnitOfMeasure>.<Name>.Value
                    NewParameter.Unit.Author = parameterItem.<UnitOfMeasure>.<Author>.Value
                    NewParameter.Unit.Code = parameterItem.<UnitOfMeasure>.<Code>.Value
                    NewTrans.ParameterValue.Add(NewParameter)
                Next

                List.Add(NewTrans)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<TransformationList>.<CreationDate>.Value
            LastEditDate = XDoc.<TransformationList>.<LastEditDate>.Value
            Description = XDoc.<TransformationList>.<Description>.Value

            Dim Transformations = From item In XDoc.<TransformationList>.<Transformation>

            List.Clear()
            For Each transItem In Transformations
                Dim NewTrans As New Transformation
                NewTrans.Name = transItem.<Name>.Value
                NewTrans.Author = transItem.<Author>.Value
                If transItem.<Code>.Value = Nothing Then
                    NewTrans.Code = 0
                Else
                    NewTrans.Code = transItem.<Code>.Value
                End If
                NewTrans.Selected = transItem.<Selected>.Value

                NewTrans.Version = transItem.<Version>.Value
                NewTrans.VariantNo = transItem.<Variant>.Value
                NewTrans.Scope = transItem.<Scope>.Value
                NewTrans.Accuracy = transItem.<Accuracy>.Value
                NewTrans.Comments = transItem.<Comments>.Value
                NewTrans.Deprecated = transItem.<Deprecated>.Value

                Dim aliasNames = From item In transItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewTrans.AddAlias(nameItem)
                Next

                'Read Source CRS information: -------------------------------------------------------------------------------------------------
                NewTrans.SourceCRS.Name = transItem.<SourceCRS>.<Name>.Value
                NewTrans.SourceCRS.Author = transItem.<SourceCRS>.<Author>.Value
                NewTrans.SourceCRS.Code = transItem.<SourceCRS>.<Code>.Value

                'Read Target CRS information: -------------------------------------------------------------------------------------------------
                NewTrans.TargetCRS.Name = transItem.<TargetCRS>.<Name>.Value
                NewTrans.TargetCRS.Author = transItem.<TargetCRS>.<Author>.Value
                NewTrans.TargetCRS.Code = transItem.<TargetCRS>.<Code>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewTrans.Area.Name = transItem.<AreaOfUse>.<Name>.Value
                NewTrans.Area.Author = transItem.<AreaOfUse>.<Author>.Value
                NewTrans.Area.Code = transItem.<AreaOfUse>.<Code>.Value

                NewTrans.ReverseOp = transItem.<ReverseOp>.Value

                'Read Method information: -------------------------------------------------------------------------------------------------
                NewTrans.Method.Name = transItem.<Method>.<Name>.Value
                NewTrans.Method.Author = transItem.<Method>.<Author>.Value
                NewTrans.Method.Code = transItem.<Method>.<Code>.Value

                'Read Source Coord Diff Unit information: -------------------------------------------------------------------------------------------------
                NewTrans.SourceCoordDiffUnit.Name = transItem.<SourceCoordDiffUnit>.<Name>.Value
                NewTrans.SourceCoordDiffUnit.Author = transItem.<SourceCoordDiffUnit>.<Author>.Value
                NewTrans.SourceCoordDiffUnit.Code = transItem.<SourceCoordDiffUnit>.<Code>.Value

                'Read Target Coord Diff Unit information: -------------------------------------------------------------------------------------------------
                NewTrans.TargetCoordDiffUnit.Name = transItem.<TargetCoordDiffUnit>.<Name>.Value
                NewTrans.TargetCoordDiffUnit.Author = transItem.<TargetCoordDiffUnit>.<Author>.Value
                NewTrans.TargetCoordDiffUnit.Code = transItem.<TargetCoordDiffUnit>.<Code>.Value

                'Read Parameter List
                Dim parameters = From item In transItem.<ParameterList>.<Parameter>

                For Each parameterItem In parameters
                    Dim NewParameter As New ValueSummary
                    NewParameter.Name = parameterItem.<Name>.Value
                    NewParameter.Value = parameterItem.<Value>.Value
                    NewParameter.Unit.Name = parameterItem.<UnitOfMeasure>.<Name>.Value
                    NewParameter.Unit.Author = parameterItem.<UnitOfMeasure>.<Author>.Value
                    NewParameter.Unit.Code = parameterItem.<UnitOfMeasure>.<Code>.Value
                    NewTrans.ParameterValue.Add(NewParameter)
                Next

                List.Add(NewTrans)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Datum list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Coordinate_Operation table into dataset ds - Only read the "transformation" records.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate_Operation] Where [COORD_OP_TYPE] = 'transformation'", myConnection)
            TableName = "CoordOp"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_CODE COORD_OP_NAME (COORD_OP_TYPE) SOURCE_CRS_CODE TARGET_CRS_CODE COORD_TFM_VERSION COORD_OP_VARIANT AREA_OF_USE_CODE COORD_OP_SCOPE COORD_OP_ACCURACY 
            'COORD_OP_METHOD_CODE UOM_CODE_SOURCE_COORD_DIFF UOM_CODE_TARGET_COORD_DIFF REMARKS DEPRECATED

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Coordinate Operation Method table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Method]", myConnection)
            TableName = "CoordOpMethod"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, FORMULA, EXAMPLE, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            'Read the Unit of Measure table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            TableName = "UnitOfMeasure"
            da.Fill(ds, TableName)
            'Unit of Measure Table fields :
            'UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            'Read the Coordinate Operation Parameter Usage table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter Usage]", myConnection)
            TableName = "CoordOpUsage"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_METHOD_CODE, PARAMETER_CODE, SORT_ORDER, PARAM_SIGN_REVERSAL

            'Read the Coordinate Operation Parameter table into dataset ds:
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter]", myConnection)
            TableName = "CoordOpParams"
            da.Fill(ds, TableName)
            'Table fields: PARAMETER_CODE, PARAMETER_NAME, DESCRIPTION, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            'Read the Coordinate Operation Parameter Value table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter Value]", myConnection)
            TableName = "CoordOpValues"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_CODE, COORD_OP_METHOD_CODE, PARAMETER_CODE, PARAMETER_VALUE, PARAM_VALUE_FILE_REF, UOM_CODE

            'Read the Coordinate Reference System table into ds
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System]", myConnection)
            TableName = "[Coordinate Reference System]"
            da.Fill(ds, TableName)

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)


            Dim expression As String

            ' variables:
            Dim NParams As Integer 'The number of parameters used to define the projection.
            Dim ParamNo As Integer 'The current parameter number.
            Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number


            Dim NRows As Integer = ds.Tables("CoordOp").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Transformation list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Transformations from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordOp table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewTrans As New Transformation
                NewTrans.Name = ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_NAME")
                NewTrans.Author = "EPSG"
                NewTrans.Code = ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_CODE")
                NewTrans.Version = ds.Tables("CoordOp").Rows(RowNo).Item("COORD_TFM_VERSION").ToString
                NewTrans.VariantNo = ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_VARIANT")
                NewTrans.Scope = ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_SCOPE").ToString
                If IsDBNull(ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_ACCURACY")) Then
                    NewTrans.Accuracy = Single.NaN
                Else
                    NewTrans.Accuracy = ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_ACCURACY")
                End If


                If IsDBNull(ds.Tables("CoordOp").Rows(RowNo).Item("REMARKS")) Then
                    NewTrans.Comments = ""
                Else
                    NewTrans.Comments = ds.Tables("CoordOp").Rows(RowNo).Item("REMARKS")
                End If

                NewTrans.Deprecated = ds.Tables("CoordOp").Rows(RowNo).Item("DEPRECATED")

                'Add Source Coordinate Reference System information -------------------------------------------------------
                NewTrans.SourceCRS.Code = ds.Tables("CoordOp").Rows(RowNo).Item("SOURCE_CRS_CODE")
                expression = "[COORD_REF_SYS_CODE] = " & Str(NewTrans.SourceCRS.Code)
                Dim sourceCRSParameters = ds.Tables("[Coordinate Reference System]").Select(expression)
                If sourceCRSParameters.Count > 0 Then
                    NewTrans.SourceCRS.Name = sourceCRSParameters(0).Item("COORD_REF_SYS_NAME")
                    NewTrans.SourceCRS.Author = "EPSG"
                End If

                'Add Target Coordinate Reference System information -------------------------------------------------------
                NewTrans.TargetCRS.Code = ds.Tables("CoordOp").Rows(RowNo).Item("TARGET_CRS_CODE")
                expression = "[COORD_REF_SYS_CODE] = " & Str(NewTrans.TargetCRS.Code)
                Dim targetCRSParameters = ds.Tables("[Coordinate Reference System]").Select(expression)
                If targetCRSParameters.Count > 0 Then
                    NewTrans.TargetCRS.Name = targetCRSParameters(0).Item("COORD_REF_SYS_NAME")
                    NewTrans.TargetCRS.Author = "EPSG"
                End If

                'Add list of alias names --------------------------------------------------------------------------------------------
                expression = "[OBJECT_TABLE_NAME] = 'Datum' AND [OBJECT_CODE] = " & Str(NewTrans.Code)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewTrans.AddAlias(item.Item("ALIAS").ToString)
                Next
                '---------------------------------------------------------------------------------------------------------------------

                'Add Area of Use information -----------------------------------------------------------------------------------------
                NewTrans.Area.Code = ds.Tables("CoordOp").Rows(RowNo).Item("AREA_OF_USE_CODE")
                expression = "[AREA_CODE] = " & Str(NewTrans.Area.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewTrans.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                    NewTrans.Area.Author = "EPSG"
                Else 'No match found. Just save the Author and Code.
                    NewTrans.Area.Author = "EPSG"
                End If

                'Add Method information -----------------------------------------------------------------------------------------
                NewTrans.Method.Code = ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_METHOD_CODE")
                expression = "[COORD_OP_METHOD_CODE] = " & Str(NewTrans.Method.Code)
                'Table fields: COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, FORMULA, EXAMPLE, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED
                Dim transMethodParameters = ds.Tables("CoordOpMethod").Select(expression)

                If transMethodParameters.Count > 0 Then
                    NewTrans.Method.Name = transMethodParameters(0).Item("COORD_OP_METHOD_NAME").ToString
                    NewTrans.Method.Author = "EPSG"
                    NewTrans.ReverseOp = transMethodParameters(0).Item("REVERSE_OP")
                Else 'No match found. Just save the Author and Code.
                    NewTrans.Method.Author = "EPSG"
                End If

                'UOM_CODE_SOURCE_COORD_DIFF  Unit of measure of the input or source coordinate differences in a polynomial operation.  Often different from the UOM of the coordinate reference system.
                If IsDBNull(ds.Tables("CoordOp").Rows(RowNo).Item("UOM_CODE_SOURCE_COORD_DIFF")) Then
                    NewTrans.SourceCoordDiffUnit.Code = 0
                    NewTrans.SourceCoordDiffUnit.Name = ""
                    NewTrans.SourceCoordDiffUnit.Author = "EPSG"
                Else
                    NewTrans.SourceCoordDiffUnit.Code = ds.Tables("CoordOp").Rows(RowNo).Item("UOM_CODE_SOURCE_COORD_DIFF")
                    expression = "[UOM_CODE] = " & Str(NewTrans.SourceCoordDiffUnit.Code)
                    Dim sourceCoordDiffUom = ds.Tables("UnitOfMeasure").Select(expression)
                    'UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED
                    If sourceCoordDiffUom.Count > 0 Then
                        NewTrans.SourceCoordDiffUnit.Name = sourceCoordDiffUom(0).Item("UNIT_OF_MEAS_NAME").ToString
                        NewTrans.SourceCoordDiffUnit.Author = "EPSG"
                    Else
                        NewTrans.SourceCoordDiffUnit.Name = ""
                        NewTrans.SourceCoordDiffUnit.Author = "EPSG"
                    End If
                End If

                'UOM_CODE_TARGET_COORD_DIFF  Unit of measure of the output or target coordinate differences in a polynomial operation.  Often different from the UOM of the coordinate reference system.
                If IsDBNull(ds.Tables("CoordOp").Rows(RowNo).Item("UOM_CODE_TARGET_COORD_DIFF")) Then
                    NewTrans.TargetCoordDiffUnit.Code = 0
                    NewTrans.TargetCoordDiffUnit.Name = ""
                    NewTrans.TargetCoordDiffUnit.Author = "EPSG"
                Else
                    NewTrans.TargetCoordDiffUnit.Code = ds.Tables("CoordOp").Rows(RowNo).Item("UOM_CODE_TARGET_COORD_DIFF")
                    expression = "[UOM_CODE] = " & Str(NewTrans.TargetCoordDiffUnit.Code)
                    Dim targetCoordDiffUom = ds.Tables("UnitOfMeasure").Select(expression)
                    'UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED
                    If targetCoordDiffUom.Count > 0 Then
                        NewTrans.TargetCoordDiffUnit.Name = targetCoordDiffUom(0).Item("UNIT_OF_MEAS_NAME").ToString
                        NewTrans.TargetCoordDiffUnit.Author = "EPSG"
                    Else
                        NewTrans.TargetCoordDiffUnit.Name = ""
                        NewTrans.TargetCoordDiffUnit.Author = "EPSG"
                    End If
                End If

                'Transformation parameters -----------------------------------------------------------------------------------------
                expression = "[COORD_OP_METHOD_CODE] = " & Str(NewTrans.Method.Code)
                Dim CoordOpUsage = ds.Tables("CoordOpUsage").Select(expression)
                'CoordOpusage Table fields: COORD_OP_METHOD_CODE, PARAMETER_CODE, SORT_ORDER, PARAM_SIGN_REVERSAL

                NParams = CoordOpUsage.Count
                For ParamNo = 0 To NParams - 1 'Process each Transformation parameter.
                    Dim NewParam As New ValueSummary
                    ParamCode = CoordOpUsage(ParamNo).Item("PARAMETER_CODE")

                    'Get the Transformation parameter details:
                    expression = "[PARAMETER_CODE] = " & Str(ParamCode)
                    Dim parameterDetails = ds.Tables("CoordOpParams").Select(expression)
                    'Table fields: PARAMETER_CODE, PARAMETER_NAME, DESCRIPTION, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

                    'Get the Transformation Parameter Value:
                    expression = "[COORD_OP_CODE] = " & Str(NewTrans.Code) & " And [PARAMETER_CODE] = " & Str(ParamCode)
                    Dim parameterValues = ds.Tables("CoordOpValues").Select(expression)
                    'Table fields: COORD_OP_CODE, COORD_OP_METHOD_CODE, PARAMETER_CODE, PARAMETER_VALUE, PARAM_VALUE_FILE_REF, UOM_CODE
                    If parameterValues.Count > 0 Then
                        If IsDBNull(parameterValues(0).Item("UOM_CODE")) Then
                            RaiseEvent ErrorMessage("Null parameter value: " & "    Transformation Name: " & ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_NAME").ToString & vbCrLf)
                            NewParam.Unit.Code = 0
                            NewParam.Unit.Name = ""
                            NewParam.Unit.Author = "EPSG"
                            NewParam.Name = parameterDetails(0).Item("PARAMETER_NAME")
                            If IsDBNull(parameterValues(0).Item("PARAMETER_VALUE")) Then
                                NewParam.Value = Double.NaN
                            Else
                                NewParam.Value = parameterValues(0).Item("PARAMETER_VALUE")
                            End If

                        Else
                            NewParam.Unit.Code = parameterValues(0).Item("UOM_CODE")
                            'Get the Unit Of Measure Details:
                            expression = "[UOM_CODE] = " & Str(NewParam.Unit.Code)
                            Dim parameterUom = ds.Tables("UnitOfMeasure").Select(expression)
                            'Table fields: UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED
                            NewParam.Name = parameterDetails(0).Item("PARAMETER_NAME")
                            If IsDBNull(parameterValues(0).Item("PARAMETER_VALUE")) Then
                                NewParam.Value = Double.NaN
                            Else
                                NewParam.Value = parameterValues(0).Item("PARAMETER_VALUE")
                            End If
                            'NewParam.Value = parameterValues(0).Item("PARAMETER_VALUE")
                            NewParam.Unit.Name = parameterUom(0).Item("UNIT_OF_MEAS_NAME")
                            NewParam.Unit.Author = "EPSG"
                        End If
                    Else
                        RaiseEvent ErrorMessage("There are no parameter values: " & expression & "    Transformation Name: " & ds.Tables("CoordOp").Rows(RowNo).Item("COORD_OP_NAME").ToString & vbCrLf)
                    End If
                    NewTrans.ParameterValue.Add(NewParam)
                Next
                List.Add(NewTrans)
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Transformations as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Transformation List File-->
                       <TransformationList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <Transformation>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Version><%= item.Version %></Version>
                                     <Variant><%= item.VariantNo %></Variant>
                                     <Scope><%= item.Scope %></Scope>
                                     <Accuracy><%= item.Accuracy %></Accuracy>
                                     <Comments><%= item.Comments %></Comments>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <SourceCRS>
                                         <Name><%= item.SourceCRS.Name %></Name>
                                         <Author><%= item.SourceCRS.Author %></Author>
                                         <Code><%= item.SourceCRS.Code %></Code>
                                     </SourceCRS>
                                     <TargetCRS>
                                         <Name><%= item.TargetCRS.Name %></Name>
                                         <Author><%= item.TargetCRS.Author %></Author>
                                         <Code><%= item.TargetCRS.Code %></Code>
                                     </TargetCRS>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                     <ReverseOp><%= item.ReverseOp %></ReverseOp>
                                     <Method>
                                         <Name><%= item.Method.Name %></Name>
                                         <Author><%= item.Method.Author %></Author>
                                         <Code><%= item.Method.Code %></Code>
                                     </Method>
                                     <SourceCoordDiffUnit>
                                         <Name><%= item.SourceCoordDiffUnit.Name %></Name>
                                         <Author><%= item.SourceCoordDiffUnit.Author %></Author>
                                         <Code><%= item.SourceCoordDiffUnit.Code %></Code>
                                     </SourceCoordDiffUnit>
                                     <TargetCoordDiffUnit>
                                         <Name><%= item.TargetCoordDiffUnit.Name %></Name>
                                         <Author><%= item.TargetCoordDiffUnit.Author %></Author>
                                         <Code><%= item.TargetCoordDiffUnit.Code %></Code>
                                     </TargetCoordDiffUnit>
                                     <ParameterList>
                                         <%= From paramItem In item.ParameterValue
                                             Select _
                                             <Parameter>
                                                 <Name><%= paramItem.Name %></Name>
                                                 <Value><%= paramItem.Value %></Value>
                                                 <UnitOfMeasure>
                                                     <Name><%= paramItem.Unit.Name %></Name>
                                                     <Author><%= paramItem.Unit.Author %></Author>
                                                     <Code><%= paramItem.Unit.Code %></Code>
                                                 </UnitOfMeasure>
                                             </Parameter>
                                         %>
                                     </ParameterList>
                                 </Transformation>
                           %>
                           <!---->
                       </TransformationList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'TransformationList ------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class GeodeticDatumList '----------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store a list of Geodetic Datum parameters.

        Public List As New List(Of GeodeticDatum)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        ''Load the XML data in the XDoc into the Geodetic Datum List.
        'Public Sub LoadXml_Old(ByRef XDoc As System.Xml.Linq.XDocument)

        '    CreationDate = XDoc.<GeodeticDatumList>.<CreationDate>.Value
        '    LastEditDate = XDoc.<GeodeticDatumList>.<LastEditDate>.Value
        '    Description = XDoc.<GeodeticDatumList>.<Description>.Value

        '    Dim GeoDatums = From item In XDoc.<GeodeticDatumList>.<GeodeticDatum>

        '    List.Clear()
        '    For Each datumItem In GeoDatums
        '        Dim NewGeoDatum As New GeodeticDatum
        '        NewGeoDatum.Name = datumItem.<Name>.Value
        '        NewGeoDatum.Author = datumItem.<Author>.Value
        '        If datumItem.<Code>.Value = Nothing Then
        '            NewGeoDatum.Code = 0
        '        Else
        '            NewGeoDatum.Code = datumItem.<Code>.Value
        '        End If
        '        NewGeoDatum.Selected = datumItem.<Selected>.Value
        '        NewGeoDatum.Comments = datumItem.<Comments>.Value
        '        NewGeoDatum.OriginDescription = datumItem.<OriginDescription>.Value
        '        NewGeoDatum.Epoch = datumItem.<RealizationEpoch>.Value
        '        NewGeoDatum.Scope = datumItem.<Scope>.Value
        '        NewGeoDatum.Deprecated = datumItem.<Deprecated>.Value

        '        Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
        '        For Each nameItem In aliasNames
        '            NewGeoDatum.AddAlias(nameItem)
        '        Next

        '        'Read Ellipsoid information: -------------------------------------------------------------------------------------------------
        '        NewGeoDatum.Ellipsoid.Name = datumItem.<Ellipsoid>.<Name>.Value
        '        NewGeoDatum.Ellipsoid.Author = datumItem.<Ellipsoid>.<Author>.Value
        '        NewGeoDatum.Ellipsoid.Code = datumItem.<Ellipsoid>.<Code>.Value
        '        NewGeoDatum.Ellipsoid.Comments = datumItem.<Ellipsoid>.<Comments>.Value
        '        Select Case datumItem.<Ellipsoid>.<EllipsoidParameters>.Value
        '            Case "SemiMajorAxis_InverseFlattening"
        '                NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_InverseFlattening
        '            Case "SemiMajorAxis_SemiMinorAxis"
        '                NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_SemiMinorAxis
        '            Case Else
        '                NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
        '        End Select

        '        'NewGeoDatum.Ellipsoid.EllipsoidParameters = datumItem.<Ellipsoid>.<EllipsoidParameters>.Value
        '        NewGeoDatum.Ellipsoid.SemiMajorAxis = datumItem.<Ellipsoid>.<SemiMajorAxis>.Value
        '        NewGeoDatum.Ellipsoid.SemiMinorAxis = datumItem.<Ellipsoid>.<SemiMinorAxis>.Value
        '        NewGeoDatum.Ellipsoid.InverseFlattening = datumItem.<Ellipsoid>.<InverseFlattening>.Value
        '        Dim aliasEllipsoidNames = From item In datumItem.<Ellipsoid>.<AliasNames>.<AliasName>
        '        For Each nameItem In aliasEllipsoidNames
        '            NewGeoDatum.Ellipsoid.AddAlias(nameItem)
        '        Next

        '        'Read Prime Meridian information: -------------------------------------------------------------------------------------------------
        '        NewGeoDatum.PrimeMeridian.Name = datumItem.<PrimeMeridian>.<Name>.Value
        '        NewGeoDatum.PrimeMeridian.Author = datumItem.<PrimeMeridian>.<Author>.Value
        '        NewGeoDatum.PrimeMeridian.Code = datumItem.<PrimeMeridian>.<Code>.Value
        '        NewGeoDatum.PrimeMeridian.Comments = datumItem.<PrimeMeridian>.<Comments>.Value
        '        NewGeoDatum.PrimeMeridian.LongitudeFromGreenwich = datumItem.<PrimeMeridian>.<LongitudeFromGreenwich>.Value
        '        Select Case datumItem.<PrimeMeridian>.<LongitudeUOM>.Value
        '            Case "Degree"
        '                NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Degree
        '            Case "Gradian"
        '                NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Gradian
        '            Case "Sexagesimal_DMS"
        '                NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Sexagesimal_DMS
        '            Case Else
        '                NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
        '        End Select
        '        'NewGeoDatum.PrimeMeridian.LongitudeUOM = datumItem.<PrimeMeridian>.<LongitudeUOM>.Value
        '        Dim aliasPmNames = From item In datumItem.<PrimeMeridian>.<AliasNames>.<AliasName>
        '        For Each nameItem In aliasPmNames
        '            NewGeoDatum.PrimeMeridian.AddAlias(nameItem)
        '        Next

        '        'Read Area of Use information: -------------------------------------------------------------------------------------------------
        '        NewGeoDatum.AreaOfUse.Name = datumItem.<AreaOfUse>.<Name>.Value
        '        NewGeoDatum.AreaOfUse.Author = datumItem.<AreaOfUse>.<Author>.Value
        '        NewGeoDatum.AreaOfUse.Code = datumItem.<AreaOfUse>.<Code>.Value
        '        NewGeoDatum.AreaOfUse.Comments = datumItem.<AreaOfUse>.<Comments>.Value
        '        NewGeoDatum.AreaOfUse.Description = datumItem.<AreaOfUse>.<Description>.Value
        '        NewGeoDatum.AreaOfUse.Deprecated = datumItem.<AreaOfUse>.<Deprecated>.Value
        '        NewGeoDatum.AreaOfUse.SouthLatitude = datumItem.<AreaOfUse>.<SouthLatitude>.Value
        '        NewGeoDatum.AreaOfUse.NorthLatitude = datumItem.<AreaOfUse>.<NorthLatitude>.Value
        '        NewGeoDatum.AreaOfUse.WestLongitude = datumItem.<AreaOfUse>.<WestLongitude>.Value
        '        NewGeoDatum.AreaOfUse.EastLongitude = datumItem.<AreaOfUse>.<EastLongitude>.Value
        '        NewGeoDatum.AreaOfUse.IsoA2Code = datumItem.<AreaOfUse>.<IsoA2Code>.Value
        '        NewGeoDatum.AreaOfUse.IsoA3Code = datumItem.<AreaOfUse>.<IsoA3Code>.Value
        '        NewGeoDatum.AreaOfUse.IsoNCode = datumItem.<AreaOfUse>.<IsoNCode>.Value
        '        Dim aliasAouNames = From item In datumItem.<AreaOfUse>.<AliasNames>.<AliasName>
        '        For Each nameItem In aliasAouNames
        '            NewGeoDatum.AreaOfUse.AddAlias(nameItem)
        '        Next

        '        List.Add(NewGeoDatum)
        '    Next
        'End Sub

        'Load the XML data in the XDoc into the Geodetic Datum List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<GeodeticDatumList>.<CreationDate>.Value
            LastEditDate = XDoc.<GeodeticDatumList>.<LastEditDate>.Value
            Description = XDoc.<GeodeticDatumList>.<Description>.Value

            Dim GeoDatums = From item In XDoc.<GeodeticDatumList>.<GeodeticDatum>

            List.Clear()
            For Each datumItem In GeoDatums
                Dim NewGeoDatum As New GeodeticDatum
                NewGeoDatum.Name = datumItem.<Name>.Value
                NewGeoDatum.Author = datumItem.<Author>.Value
                If datumItem.<Code>.Value = Nothing Then
                    NewGeoDatum.Code = 0
                Else
                    NewGeoDatum.Code = datumItem.<Code>.Value
                End If
                NewGeoDatum.Selected = datumItem.<Selected>.Value
                NewGeoDatum.Comments = datumItem.<Comments>.Value
                NewGeoDatum.OriginDescription = datumItem.<OriginDescription>.Value
                NewGeoDatum.Epoch = datumItem.<RealizationEpoch>.Value
                NewGeoDatum.Scope = datumItem.<Scope>.Value
                NewGeoDatum.Deprecated = datumItem.<Deprecated>.Value

                Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewGeoDatum.AddAlias(nameItem)
                Next

                'Read Ellipsoid information: -------------------------------------------------------------------------------------------------
                NewGeoDatum.Ellipsoid.Name = datumItem.<Ellipsoid>.<Name>.Value
                NewGeoDatum.Ellipsoid.Author = datumItem.<Ellipsoid>.<Author>.Value
                NewGeoDatum.Ellipsoid.Code = datumItem.<Ellipsoid>.<Code>.Value
                'NewGeoDatum.Ellipsoid.Comments = datumItem.<Ellipsoid>.<Comments>.Value
                'Select Case datumItem.<Ellipsoid>.<EllipsoidParameters>.Value
                '    Case "SemiMajorAxis_InverseFlattening"
                '        NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_InverseFlattening
                '    Case "SemiMajorAxis_SemiMinorAxis"
                '        NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_SemiMinorAxis
                '    Case Else
                '        NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
                'End Select

                'NewGeoDatum.Ellipsoid.SemiMajorAxis = datumItem.<Ellipsoid>.<SemiMajorAxis>.Value
                'NewGeoDatum.Ellipsoid.SemiMinorAxis = datumItem.<Ellipsoid>.<SemiMinorAxis>.Value
                'NewGeoDatum.Ellipsoid.InverseFlattening = datumItem.<Ellipsoid>.<InverseFlattening>.Value
                'Dim aliasEllipsoidNames = From item In datumItem.<Ellipsoid>.<AliasNames>.<AliasName>
                'For Each nameItem In aliasEllipsoidNames
                '    NewGeoDatum.Ellipsoid.AddAlias(nameItem)
                'Next

                'Read Prime Meridian information: -------------------------------------------------------------------------------------------------
                NewGeoDatum.PrimeMeridian.Name = datumItem.<PrimeMeridian>.<Name>.Value
                NewGeoDatum.PrimeMeridian.Author = datumItem.<PrimeMeridian>.<Author>.Value
                NewGeoDatum.PrimeMeridian.Code = datumItem.<PrimeMeridian>.<Code>.Value
                'NewGeoDatum.PrimeMeridian.Comments = datumItem.<PrimeMeridian>.<Comments>.Value
                'NewGeoDatum.PrimeMeridian.LongitudeFromGreenwich = datumItem.<PrimeMeridian>.<LongitudeFromGreenwich>.Value
                'Select Case datumItem.<PrimeMeridian>.<LongitudeUOM>.Value
                '    Case "Degree"
                '        NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Degree
                '    Case "Gradian"
                '        NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Gradian
                '    Case "Sexagesimal_DMS"
                '        NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Sexagesimal_DMS
                '    Case Else
                '        NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
                'End Select
                ''NewGeoDatum.PrimeMeridian.LongitudeUOM = datumItem.<PrimeMeridian>.<LongitudeUOM>.Value
                'Dim aliasPmNames = From item In datumItem.<PrimeMeridian>.<AliasNames>.<AliasName>
                'For Each nameItem In aliasPmNames
                '    NewGeoDatum.PrimeMeridian.AddAlias(nameItem)
                'Next

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewGeoDatum.Area.Name = datumItem.<AreaOfUse>.<Name>.Value
                NewGeoDatum.Area.Author = datumItem.<AreaOfUse>.<Author>.Value
                NewGeoDatum.Area.Code = datumItem.<AreaOfUse>.<Code>.Value
                'NewGeoDatum.AreaOfUse.Comments = datumItem.<AreaOfUse>.<Comments>.Value
                'NewGeoDatum.AreaOfUse.Description = datumItem.<AreaOfUse>.<Description>.Value
                'NewGeoDatum.AreaOfUse.Deprecated = datumItem.<AreaOfUse>.<Deprecated>.Value
                'NewGeoDatum.AreaOfUse.SouthLatitude = datumItem.<AreaOfUse>.<SouthLatitude>.Value
                'NewGeoDatum.AreaOfUse.NorthLatitude = datumItem.<AreaOfUse>.<NorthLatitude>.Value
                'NewGeoDatum.AreaOfUse.WestLongitude = datumItem.<AreaOfUse>.<WestLongitude>.Value
                'NewGeoDatum.AreaOfUse.EastLongitude = datumItem.<AreaOfUse>.<EastLongitude>.Value
                'NewGeoDatum.AreaOfUse.IsoA2Code = datumItem.<AreaOfUse>.<IsoA2Code>.Value
                'NewGeoDatum.AreaOfUse.IsoA3Code = datumItem.<AreaOfUse>.<IsoA3Code>.Value
                'NewGeoDatum.AreaOfUse.IsoNCode = datumItem.<AreaOfUse>.<IsoNCode>.Value
                'Dim aliasAouNames = From item In datumItem.<AreaOfUse>.<AliasNames>.<AliasName>
                'For Each nameItem In aliasAouNames
                '    NewGeoDatum.AreaOfUse.AddAlias(nameItem)
                'Next

                List.Add(NewGeoDatum)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<GeodeticDatumList>.<CreationDate>.Value
            LastEditDate = XDoc.<GeodeticDatumList>.<LastEditDate>.Value
            Description = XDoc.<GeodeticDatumList>.<Description>.Value

            Dim GeoDatums = From item In XDoc.<GeodeticDatumList>.<GeodeticDatum>

            List.Clear()
            For Each datumItem In GeoDatums
                Dim NewGeoDatum As New GeodeticDatum
                NewGeoDatum.Name = datumItem.<Name>.Value
                NewGeoDatum.Author = datumItem.<Author>.Value
                If datumItem.<Code>.Value = Nothing Then
                    NewGeoDatum.Code = 0
                Else
                    NewGeoDatum.Code = datumItem.<Code>.Value
                End If
                NewGeoDatum.Selected = datumItem.<Selected>.Value
                NewGeoDatum.Comments = datumItem.<Comments>.Value
                NewGeoDatum.OriginDescription = datumItem.<OriginDescription>.Value
                NewGeoDatum.Epoch = datumItem.<RealizationEpoch>.Value
                NewGeoDatum.Scope = datumItem.<Scope>.Value
                NewGeoDatum.Deprecated = datumItem.<Deprecated>.Value

                Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewGeoDatum.AddAlias(nameItem)
                Next

                'Read Ellipsoid information: -------------------------------------------------------------------------------------------------
                NewGeoDatum.Ellipsoid.Name = datumItem.<Ellipsoid>.<Name>.Value
                NewGeoDatum.Ellipsoid.Author = datumItem.<Ellipsoid>.<Author>.Value
                NewGeoDatum.Ellipsoid.Code = datumItem.<Ellipsoid>.<Code>.Value
                'NewGeoDatum.Ellipsoid.Comments = datumItem.<Ellipsoid>.<Comments>.Value
                'Select Case datumItem.<Ellipsoid>.<EllipsoidParameters>.Value
                '    Case "SemiMajorAxis_InverseFlattening"
                '        NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_InverseFlattening
                '    Case "SemiMajorAxis_SemiMinorAxis"
                '        NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_SemiMinorAxis
                '    Case Else
                '        NewGeoDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
                'End Select

                ''NewGeoDatum.Ellipsoid.EllipsoidParameters = datumItem.<Ellipsoid>.<EllipsoidParameters>.Value
                'NewGeoDatum.Ellipsoid.SemiMajorAxis = datumItem.<Ellipsoid>.<SemiMajorAxis>.Value
                'NewGeoDatum.Ellipsoid.SemiMinorAxis = datumItem.<Ellipsoid>.<SemiMinorAxis>.Value
                'NewGeoDatum.Ellipsoid.InverseFlattening = datumItem.<Ellipsoid>.<InverseFlattening>.Value
                'Dim aliasEllipsoidNames = From item In datumItem.<Ellipsoid>.<AliasNames>.<AliasName>
                'For Each nameItem In aliasEllipsoidNames
                '    NewGeoDatum.Ellipsoid.AddAlias(nameItem)
                'Next

                'Read Prime Meridian information: -------------------------------------------------------------------------------------------------
                NewGeoDatum.PrimeMeridian.Name = datumItem.<PrimeMeridian>.<Name>.Value
                NewGeoDatum.PrimeMeridian.Author = datumItem.<PrimeMeridian>.<Author>.Value
                NewGeoDatum.PrimeMeridian.Code = datumItem.<PrimeMeridian>.<Code>.Value
                'NewGeoDatum.PrimeMeridian.Comments = datumItem.<PrimeMeridian>.<Comments>.Value
                'NewGeoDatum.PrimeMeridian.LongitudeFromGreenwich = datumItem.<PrimeMeridian>.<LongitudeFromGreenwich>.Value
                'Select Case datumItem.<PrimeMeridian>.<LongitudeUOM>.Value
                '    Case "Degree"
                '        NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Degree
                '    Case "Gradian"
                '        NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Gradian
                '    Case "Sexagesimal_DMS"
                '        NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Sexagesimal_DMS
                '    Case Else
                '        NewGeoDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
                'End Select
                ''NewGeoDatum.PrimeMeridian.LongitudeUOM = datumItem.<PrimeMeridian>.<LongitudeUOM>.Value
                'Dim aliasPmNames = From item In datumItem.<PrimeMeridian>.<AliasNames>.<AliasName>
                'For Each nameItem In aliasPmNames
                '    NewGeoDatum.PrimeMeridian.AddAlias(nameItem)
                'Next

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewGeoDatum.Area.Name = datumItem.<AreaOfUse>.<Name>.Value
                NewGeoDatum.Area.Author = datumItem.<AreaOfUse>.<Author>.Value
                NewGeoDatum.Area.Code = datumItem.<AreaOfUse>.<Code>.Value
                'NewGeoDatum.AreaOfUse.Comments = datumItem.<AreaOfUse>.<Comments>.Value
                'NewGeoDatum.AreaOfUse.Description = datumItem.<AreaOfUse>.<Description>.Value
                'NewGeoDatum.AreaOfUse.Deprecated = datumItem.<AreaOfUse>.<Deprecated>.Value
                'NewGeoDatum.AreaOfUse.SouthLatitude = datumItem.<AreaOfUse>.<SouthLatitude>.Value
                'NewGeoDatum.AreaOfUse.NorthLatitude = datumItem.<AreaOfUse>.<NorthLatitude>.Value
                'NewGeoDatum.AreaOfUse.WestLongitude = datumItem.<AreaOfUse>.<WestLongitude>.Value
                'NewGeoDatum.AreaOfUse.EastLongitude = datumItem.<AreaOfUse>.<EastLongitude>.Value
                'NewGeoDatum.AreaOfUse.IsoA2Code = datumItem.<AreaOfUse>.<IsoA2Code>.Value
                'NewGeoDatum.AreaOfUse.IsoA3Code = datumItem.<AreaOfUse>.<IsoA3Code>.Value
                'NewGeoDatum.AreaOfUse.IsoNCode = datumItem.<AreaOfUse>.<IsoNCode>.Value
                'Dim aliasAouNames = From item In datumItem.<AreaOfUse>.<AliasNames>.<AliasName>
                'For Each nameItem In aliasAouNames
                '    NewGeoDatum.AreaOfUse.AddAlias(nameItem)
                'Next

                List.Add(NewGeoDatum)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Datum list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()


            'Read the Datum table into dataset ds - Only read the geodetic datum records.
            da = New OleDb.OleDbDataAdapter("Select * From [Datum] Where [DATUM_TYPE] = 'geodetic'", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)

            'Read the Ellipsoid table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Ellipsoid]", myConnection)
            TableName = "Ellipsoid"
            da.Fill(ds, TableName)

            'Read the Prime Meridian table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Prime Meridian]", myConnection)
            TableName = "Prime Meridian"
            da.Fill(ds, TableName)

            'Read the Unit of Measure table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            TableName = "Unit of Measure"
            da.Fill(ds, TableName)
            'Unit of Measure Table fields :
            'UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            Dim UOMCode As Integer


            Dim NRows As Integer = ds.Tables("Datum").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Geodetic Datum list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            'Dim EllipsoidParametersString As String
            RaiseEvent Message("Reading Geodetic Datums from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the Datum table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewDatum As New GeodeticDatum
                NewDatum.Name = ds.Tables("Datum").Rows(RowNo).Item("DATUM_NAME")
                NewDatum.Author = "EPSG"
                NewDatum.Code = ds.Tables("Datum").Rows(RowNo).Item("DATUM_CODE")
                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REMARKS")) Then
                    NewDatum.Comments = ""
                Else
                    NewDatum.Comments = ds.Tables("Datum").Rows(RowNo).Item("REMARKS")
                End If

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")) Then
                    NewDatum.OriginDescription = ""
                Else
                    NewDatum.OriginDescription = ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")
                End If

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")) Then
                    NewDatum.Epoch = ""
                Else
                    NewDatum.Epoch = ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")
                End If

                NewDatum.Scope = ds.Tables("Datum").Rows(RowNo).Item("DATUM_SCOPE")
                NewDatum.Deprecated = ds.Tables("Datum").Rows(RowNo).Item("DEPRECATED")

                'TO DO: CHECK OBJECT_TABLE_NAME ANBD OBJECT_CODE FOR EVERY LoadEpsgDbList METHOD
                'Add list of alias names --------------------------------------------------------------------------------------------
                expression = "[OBJECT_TABLE_NAME] = 'Datum' AND [OBJECT_CODE] = " & Str(NewDatum.Code)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewDatum.AddAlias(item.Item("ALIAS").ToString)
                Next
                '---------------------------------------------------------------------------------------------------------------------

                'Add ellipsoid parameters -----------------------------------------------------------------------------------
                'In this updated version of LoadEpsgDbList, only the Ellipsoids Name, Author and Code are added to each Geodetic Datum record
                NewDatum.Ellipsoid.Code = ds.Tables("Datum").Rows(RowNo).Item("ELLIPSOID_CODE")
                NewDatum.Ellipsoid.Author = "EPSG"
                expression = "[ELLIPSOID_CODE] = " & Str(NewDatum.Ellipsoid.Code)
                Dim ellipsoidParameters = ds.Tables("Ellipsoid").Select(expression)
                If ellipsoidParameters.Count > 0 Then
                    NewDatum.Ellipsoid.Name = ellipsoidParameters(0).Item("ELLIPSOID_NAME").ToString
                    'NewDatum.Ellipsoid.Comments = ellipsoidParameters(0).Item("REMARKS").ToString

                    'If IsDBNull(ellipsoidParameters(0).Item("INV_FLATTENING")) Then
                    '    If IsDBNull(ellipsoidParameters(0).Item("SEMI_MINOR_AXIS")) Then
                    '        'EllipsoidParametersString = "Unknown"
                    '        NewDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
                    '    Else
                    '        'EllipsoidParametersString = "SemiMajorAxis_SemiMinorAxis"
                    '        NewDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_SemiMinorAxis
                    '    End If
                    '    'EllipsoidParametersString = "SemiMajorAxis_SemiMinorAxis"
                    'Else 'Inverse Flattening parameter defined
                    '    If IsDBNull(ellipsoidParameters(0).Item("SEMI_MINOR_AXIS")) Then
                    '        'EllipsoidParametersString = "SemiMajorAxis_InverseFlattening"
                    '        NewDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.SemiMajorAxis_InverseFlattening
                    '    Else
                    '        'EllipsoidParametersString = "Unknown"
                    '        NewDatum.Ellipsoid.EllipsoidParameters = Ellipsoid.DefiningParameters.Unknown
                    '    End If
                    'End If

                    'NewDatum.Ellipsoid.SemiMajorAxis = ellipsoidParameters(0).Item("SEMI_MAJOR_AXIS")
                    'If IsDBNull(ellipsoidParameters(0).Item("INV_FLATTENING")) Then
                    '    NewDatum.Ellipsoid.InverseFlattening = Double.NaN
                    'Else
                    '    NewDatum.Ellipsoid.InverseFlattening = ellipsoidParameters(0).Item("INV_FLATTENING")
                    'End If

                    'If IsDBNull(ellipsoidParameters(0).Item("SEMI_MINOR_AXIS")) Then
                    '    NewDatum.Ellipsoid.SemiMinorAxis = Double.NaN
                    'Else
                    '    NewDatum.Ellipsoid.SemiMinorAxis = ellipsoidParameters(0).Item("SEMI_MINOR_AXIS")
                    'End If


                    ''Add list of alias names --------------------------------------------------------------------------------------------
                    'expression = "[OBJECT_TABLE_NAME] = 'Ellipsoid' AND [OBJECT_CODE] = " & Str(NewDatum.Ellipsoid.Code)
                    'Dim result2 = ds.Tables("Alias").Select(expression)
                    'For Each item In result2
                    '    NewDatum.Ellipsoid.AddAlias(item.Item("ALIAS").ToString)
                    'Next
                Else
                    'No ellipsoid parameters found
                End If

                'Add prime meridian parameters -----------------------------------------------------------------------------------
                NewDatum.PrimeMeridian.Code = ds.Tables("Datum").Rows(RowNo).Item("PRIME_MERIDIAN_CODE")
                NewDatum.PrimeMeridian.Author = "EPSG"
                expression = "[PRIME_MERIDIAN_CODE] = " & Str(NewDatum.PrimeMeridian.Code)
                Dim primeMeridianParameters = ds.Tables("Prime Meridian").Select(expression)
                If primeMeridianParameters.Count > 0 Then
                    NewDatum.PrimeMeridian.Name = primeMeridianParameters(0).Item("PRIME_MERIDIAN_NAME").ToString
                    'NewDatum.PrimeMeridian.Author = "EPSG"
                    'NewDatum.PrimeMeridian.Comments = primeMeridianParameters(0).Item("REMARKS").ToString
                    'NewDatum.PrimeMeridian.LongitudeFromGreenwich = primeMeridianParameters(0).Item("GREENWICH_LONGITUDE")

                    'UOMCode = primeMeridianParameters(0).Item("UOM_CODE")
                    'expression = "[UOM_CODE] = " & Str(UOMCode)
                    'Dim UomResult = ds.Tables("Unit of Measure").Select(expression)
                    'If UomResult.Count > 0 Then
                    '    Select Case UomResult(0).Item("UNIT_OF_MEAS_NAME").ToString
                    '        Case "degree"
                    '            NewDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Degree
                    '        Case "grad"
                    '            NewDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Gradian
                    '        Case "sexagesimal DMS"
                    '            NewDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Sexagesimal_DMS
                    '        Case Else
                    '            NewDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
                    '    End Select
                    'Else
                    '    NewDatum.PrimeMeridian.LongitudeUOM = PrimeMeridian.LongitudeUnits.Unknown
                    'End If

                    ''Add list of alias names --------------------------------------------------------------------------------------------
                    'expression = "[OBJECT_TABLE_NAME] = 'Prime Meridian' AND [OBJECT_CODE] = " & Str(NewDatum.PrimeMeridian.Code)
                    'Dim pmAliasResult = ds.Tables("Alias").Select(expression)
                    'For Each item In pmAliasResult
                    '    NewDatum.PrimeMeridian.AddAlias(item.Item("ALIAS").ToString)
                    'Next
                Else
                    'No prime meridian perameters found.
                End If

                'Add Area of Use information -----------------------------------------------------------------------------------------
                NewDatum.Area.Code = ds.Tables("Datum").Rows(RowNo).Item("AREA_OF_USE_CODE")
                NewDatum.Area.Author = "EPSG"
                expression = "[AREA_CODE] = " & Str(NewDatum.Area.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewDatum.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                    'NewDatum.AreaOfUse.Author = "EPSG"
                    ' NewDatum.AreaOfUse.Description = areaOfUseParameters(0).Item("AREA_OF_USE").ToString
                    'NewDatum.AreaOfUse.Comments = areaOfUseParameters(0).Item("REMARKS").ToString
                    'If IsDBNull(areaOfUseParameters(0).Item("AREA_SOUTH_BOUND_LAT")) Then
                    '    NewDatum.AreaOfUse.SouthLatitude = Double.NaN
                    'Else
                    '    NewDatum.AreaOfUse.SouthLatitude = areaOfUseParameters(0).Item("AREA_SOUTH_BOUND_LAT")
                    'End If
                    'If IsDBNull(areaOfUseParameters(0).Item("AREA_NORTH_BOUND_LAT")) Then
                    '    NewDatum.AreaOfUse.NorthLatitude = Double.NaN
                    'Else
                    '    NewDatum.AreaOfUse.NorthLatitude = areaOfUseParameters(0).Item("AREA_NORTH_BOUND_LAT")
                    'End If
                    'If IsDBNull(areaOfUseParameters(0).Item("AREA_WEST_BOUND_LON")) Then
                    '    NewDatum.AreaOfUse.WestLongitude = Double.NaN
                    'Else
                    '    NewDatum.AreaOfUse.WestLongitude = areaOfUseParameters(0).Item("AREA_WEST_BOUND_LON")
                    'End If
                    'If IsDBNull(areaOfUseParameters(0).Item("AREA_EAST_BOUND_LON")) Then
                    '    NewDatum.AreaOfUse.EastLongitude = Double.NaN
                    'Else
                    '    NewDatum.AreaOfUse.EastLongitude = areaOfUseParameters(0).Item("AREA_EAST_BOUND_LON")
                    'End If

                    'If IsDBNull(areaOfUseParameters(0).Item("ISO_A2_CODE")) Then
                    '    NewDatum.AreaOfUse.IsoA2Code = 0
                    'Else
                    '    NewDatum.AreaOfUse.IsoA2Code = areaOfUseParameters(0).Item("ISO_A2_CODE")
                    'End If
                    'If IsDBNull(areaOfUseParameters(0).Item("ISO_A3_CODE")) Then
                    '    NewDatum.AreaOfUse.IsoA3Code = 0
                    'Else
                    '    NewDatum.AreaOfUse.IsoA3Code = areaOfUseParameters(0).Item("ISO_A3_CODE")
                    'End If
                    'If IsDBNull(areaOfUseParameters(0).Item("ISO_N_CODE")) Then
                    '    NewDatum.AreaOfUse.IsoNCode = 0
                    'Else
                    '    NewDatum.AreaOfUse.IsoNCode = areaOfUseParameters(0).Item("ISO_N_CODE")
                    'End If

                    ''Add list of alias names --------------------------------------------------------------------------------------------
                    'Dim aouAliasNames As New XElement("AliasNames")
                    'expression = "[OBJECT_TABLE_NAME] = 'Area' AND [OBJECT_CODE] = " & Str(NewDatum.AreaOfUse.Code)
                    'Dim result3 = ds.Tables("Alias").Select(expression)
                    'For Each item In result3
                    '    NewDatum.AreaOfUse.AliasName.Add(item.Item("ALIAS").ToString)
                    'Next
                End If
                List.Add(NewDatum)
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        ''Function to return the list of Geodetic Datums as an XDocument
        'Public Function ToXDoc_Old() As System.Xml.Linq.XDocument

        '    Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
        '               <!---->
        '               <!--Geodetic Datum List File-->
        '               <GeodeticDatumList>
        '                   <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
        '                   <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
        '                   <Description><%= Description %></Description>
        '                   <!---->
        '                   <%= From item In List _
        '                           Select _
        '                         <GeodeticDatum>
        '                             <Name><%= item.Name %></Name>
        '                             <Author><%= item.Author %></Author>
        '                             <Code><%= item.Code %></Code>
        '                             <Selected><%= item.Selected %></Selected>
        '                             <Comments><%= item.Comments %></Comments>
        '                             <OriginDescription><%= item.OriginDescription %></OriginDescription>
        '                             <RealizationEpoch><%= item.Epoch %></RealizationEpoch>
        '                             <Scope><%= item.Scope %></Scope>
        '                             <Deprecated><%= item.Deprecated %></Deprecated>
        '                             <AliasNames>
        '                                 <%= From nameItem In item.AliasName _
        '                                     Select _
        '                                     <AliasName><%= nameItem %></AliasName>
        '                                 %>
        '                             </AliasNames>
        '                             <Ellipsoid>
        '                                 <Name><%= item.Ellipsoid.Name %></Name>
        '                                 <Author><%= item.Ellipsoid.Author %></Author>
        '                                 <Code><%= item.Ellipsoid.Code %></Code>
        '                                 <Comments><%= item.Ellipsoid.Comments %></Comments>
        '                                 <EllipsoidParameters><%= item.Ellipsoid.EllipsoidParameters %></EllipsoidParameters>
        '                                 <SemiMajorAxis><%= item.Ellipsoid.SemiMajorAxis %></SemiMajorAxis>
        '                                 <SemiMinorAxis><%= item.Ellipsoid.SemiMinorAxis %></SemiMinorAxis>
        '                                 <InverseFlattening><%= item.Ellipsoid.InverseFlattening %></InverseFlattening>
        '                                 <AliasNames>
        '                                     <%= From nameItem In item.Ellipsoid.AliasName _
        '                                     Select _
        '                                     <AliasName><%= nameItem %></AliasName>
        '                                     %>
        '                                 </AliasNames>
        '                             </Ellipsoid>
        '                             <PrimeMeridian>
        '                                 <Name><%= item.PrimeMeridian.Name %></Name>
        '                                 <Author><%= item.PrimeMeridian.Author %></Author>
        '                                 <Code><%= item.PrimeMeridian.Code %></Code>
        '                                 <Comments><%= item.PrimeMeridian.Comments %></Comments>
        '                                 <LongitudeFromGreenwich><%= item.PrimeMeridian.LongitudeFromGreenwich %></LongitudeFromGreenwich>
        '                                 <LongitudeUOM><%= item.PrimeMeridian.LongitudeUOM %></LongitudeUOM>
        '                                 <AliasNames>
        '                                     <%= From nameItem In item.PrimeMeridian.AliasName _
        '                                     Select _
        '                                     <AliasName><%= nameItem %></AliasName>
        '                                     %>
        '                                 </AliasNames>
        '                             </PrimeMeridian>
        '                             <AreaOfUse>
        '                                 <Name><%= item.AreaOfUse.Name %></Name>
        '                                 <Author><%= item.AreaOfUse.Author %></Author>
        '                                 <Code><%= item.AreaOfUse.Code %></Code>
        '                                 <Comments><%= item.AreaOfUse.Comments %></Comments>
        '                                 <Description><%= item.AreaOfUse.Description %></Description>
        '                                 <Deprecated><%= item.AreaOfUse.Deprecated %></Deprecated>
        '                                 <SouthLatitude><%= item.AreaOfUse.SouthLatitude %></SouthLatitude>
        '                                 <NorthLatitude><%= item.AreaOfUse.NorthLatitude %></NorthLatitude>
        '                                 <WestLongitude><%= item.AreaOfUse.WestLongitude %></WestLongitude>
        '                                 <EastLongitude><%= item.AreaOfUse.EastLongitude %></EastLongitude>
        '                                 <IsoA2Code><%= item.AreaOfUse.IsoA2Code %></IsoA2Code>
        '                                 <IsoA3Code><%= item.AreaOfUse.IsoA3Code %></IsoA3Code>
        '                                 <IsoNCode><%= item.AreaOfUse.IsoNCode %></IsoNCode>
        '                                 <AliasNames>
        '                                     <%= From nameItem In item.AreaOfUse.AliasName _
        '                                     Select _
        '                                     <AliasName><%= nameItem %></AliasName>
        '                                     %>
        '                                 </AliasNames>
        '                             </AreaOfUse>
        '                         </GeodeticDatum>
        '                   %>
        '                   <!---->
        '               </GeodeticDatumList>
        '    Return XDoc
        'End Function

        'Function to return the list of Geodetic Datums as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Geodetic Datum List File-->
                       <GeodeticDatumList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <GeodeticDatum>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <OriginDescription><%= item.OriginDescription %></OriginDescription>
                                     <RealizationEpoch><%= item.Epoch %></RealizationEpoch>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <Ellipsoid>
                                         <Name><%= item.Ellipsoid.Name %></Name>
                                         <Author><%= item.Ellipsoid.Author %></Author>
                                         <Code><%= item.Ellipsoid.Code %></Code>
                                     </Ellipsoid>
                                     <PrimeMeridian>
                                         <Name><%= item.PrimeMeridian.Name %></Name>
                                         <Author><%= item.PrimeMeridian.Author %></Author>
                                         <Code><%= item.PrimeMeridian.Code %></Code>
                                     </PrimeMeridian>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </GeodeticDatum>
                           %>
                           <!---->
                       </GeodeticDatumList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'GeodeticDatumList -------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class VerticalDatumList '----------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store a list of Vertical Datum parameters.

        Public List As New List(Of VerticalDatum)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Geodetic Datum List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<VerticalDatumList>.<CreationDate>.Value
            LastEditDate = XDoc.<VerticalDatumList>.<LastEditDate>.Value
            Description = XDoc.<VerticalDatumList>.<Description>.Value

            Dim VerticalDatums = From item In XDoc.<VerticalDatumList>.<VerticalDatum>

            List.Clear()
            For Each datumItem In VerticalDatums
                Dim NewVerticalDatum As New VerticalDatum
                NewVerticalDatum.Name = datumItem.<Name>.Value
                NewVerticalDatum.Author = datumItem.<Author>.Value
                If datumItem.<Code>.Value = Nothing Then
                    NewVerticalDatum.Code = 0
                Else
                    NewVerticalDatum.Code = datumItem.<Code>.Value
                End If
                NewVerticalDatum.Selected = datumItem.<Selected>.Value
                NewVerticalDatum.Comments = datumItem.<Comments>.Value
                NewVerticalDatum.OriginDescription = datumItem.<OriginDescription>.Value
                NewVerticalDatum.Epoch = datumItem.<RealizationEpoch>.Value
                NewVerticalDatum.Scope = datumItem.<Scope>.Value
                NewVerticalDatum.Deprecated = datumItem.<Deprecated>.Value

                Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewVerticalDatum.AddAlias(nameItem)
                Next

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewVerticalDatum.Area.Name = datumItem.<AreaOfUse>.<Name>.Value
                NewVerticalDatum.Area.Author = datumItem.<AreaOfUse>.<Author>.Value
                NewVerticalDatum.Area.Code = datumItem.<AreaOfUse>.<Code>.Value

                List.Add(NewVerticalDatum)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<VerticalDatumList>.<CreationDate>.Value
            LastEditDate = XDoc.<VerticalDatumList>.<LastEditDate>.Value
            Description = XDoc.<VerticalDatumList>.<Description>.Value

            Dim VerticalDatums = From item In XDoc.<VerticalDatumList>.<VerticalDatum>

            List.Clear()
            For Each datumItem In VerticalDatums
                Dim NewVerticalDatum As New VerticalDatum
                NewVerticalDatum.Name = datumItem.<Name>.Value
                NewVerticalDatum.Author = datumItem.<Author>.Value
                If datumItem.<Code>.Value = Nothing Then
                    NewVerticalDatum.Code = 0
                Else
                    NewVerticalDatum.Code = datumItem.<Code>.Value
                End If
                NewVerticalDatum.Selected = datumItem.<Selected>.Value
                NewVerticalDatum.Comments = datumItem.<Comments>.Value
                NewVerticalDatum.OriginDescription = datumItem.<OriginDescription>.Value
                NewVerticalDatum.Epoch = datumItem.<RealizationEpoch>.Value
                NewVerticalDatum.Scope = datumItem.<Scope>.Value
                NewVerticalDatum.Deprecated = datumItem.<Deprecated>.Value

                Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewVerticalDatum.AddAlias(nameItem)
                Next

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewVerticalDatum.Area.Name = datumItem.<AreaOfUse>.<Name>.Value
                NewVerticalDatum.Area.Author = datumItem.<AreaOfUse>.<Author>.Value
                NewVerticalDatum.Area.Code = datumItem.<AreaOfUse>.<Code>.Value
         
                List.Add(NewVerticalDatum)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Datum list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Datum table into dataset ds - Only read the vertical datum records.
            da = New OleDb.OleDbDataAdapter("Select * From [Datum] Where [DATUM_TYPE] = 'vertical'", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)

            ''Read the Ellipsoid table into dataset ds
            'da.SelectCommand = New OleDb.OleDbCommand("Select * From [Ellipsoid]", myConnection)
            'TableName = "Ellipsoid"
            'da.Fill(ds, TableName)

            ''Read the Prime Meridian table into dataset ds
            'da.SelectCommand = New OleDb.OleDbCommand("Select * From [Prime Meridian]", myConnection)
            'TableName = "Prime Meridian"
            'da.Fill(ds, TableName)

            ''Read the Unit of Measure table into ds
            'da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            'TableName = "Unit of Measure"
            'da.Fill(ds, TableName)
            ''Unit of Measure Table fields :
            ''UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            Dim UOMCode As Integer

            Dim NRows As Integer = ds.Tables("Datum").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Vertical Datum list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Vertical Datums from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the Datum table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewDatum As New VerticalDatum
                NewDatum.Name = ds.Tables("Datum").Rows(RowNo).Item("DATUM_NAME")
                NewDatum.Author = "EPSG"
                NewDatum.Code = ds.Tables("Datum").Rows(RowNo).Item("DATUM_CODE")
                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REMARKS")) Then
                    NewDatum.Comments = ""
                Else
                    NewDatum.Comments = ds.Tables("Datum").Rows(RowNo).Item("REMARKS")
                End If

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")) Then
                    NewDatum.OriginDescription = ""
                Else
                    NewDatum.OriginDescription = ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")
                End If

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")) Then
                    NewDatum.Epoch = ""
                Else
                    NewDatum.Epoch = ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")
                End If

                NewDatum.Scope = ds.Tables("Datum").Rows(RowNo).Item("DATUM_SCOPE")
                NewDatum.Deprecated = ds.Tables("Datum").Rows(RowNo).Item("DEPRECATED")

                'TO DO: CHECK OBJECT_TABLE_NAME AND OBJECT_CODE FOR EVERY LoadEpsgDbList METHOD
                'Add list of alias names --------------------------------------------------------------------------------------------
                expression = "[OBJECT_TABLE_NAME] = 'Datum' AND [OBJECT_CODE] = " & Str(NewDatum.Code)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewDatum.AddAlias(item.Item("ALIAS").ToString)
                Next
                '---------------------------------------------------------------------------------------------------------------------

                'Add Area of Use information -----------------------------------------------------------------------------------------
                NewDatum.Area.Code = ds.Tables("Datum").Rows(RowNo).Item("AREA_OF_USE_CODE")
                NewDatum.Area.Author = "EPSG"
                expression = "[AREA_CODE] = " & Str(NewDatum.Area.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewDatum.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                End If
                List.Add(NewDatum)
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Vertical Datums as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Vertical Datum List File-->
                       <VerticalDatumList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <VerticalDatum>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <OriginDescription><%= item.OriginDescription %></OriginDescription>
                                     <RealizationEpoch><%= item.Epoch %></RealizationEpoch>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </VerticalDatum>
                           %>
                           <!---->
                       </VerticalDatumList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'VerticalDatumList -------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class EngineeringDatumList '-------------------------------------------------------------------------------------------------------------------------------------------------------
        'Class used to store a list of Engineering Datum parameters.

        Public List As New List(Of EngineeringDatum)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Geodetic Datum List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<EngineeringDatumList>.<CreationDate>.Value
            LastEditDate = XDoc.<EngineeringDatumList>.<LastEditDate>.Value
            Description = XDoc.<EngineeringDatumList>.<Description>.Value

            Dim EngineeringDatums = From item In XDoc.<EngineeringDatumList>.<EngineeringDatum>

            List.Clear()
            For Each datumItem In EngineeringDatums
                Dim NewEngineeringDatum As New EngineeringDatum
                NewEngineeringDatum.Name = datumItem.<Name>.Value
                NewEngineeringDatum.Author = datumItem.<Author>.Value
                If datumItem.<Code>.Value = Nothing Then
                    NewEngineeringDatum.Code = 0
                Else
                    NewEngineeringDatum.Code = datumItem.<Code>.Value
                End If
                NewEngineeringDatum.Selected = datumItem.<Selected>.Value
                NewEngineeringDatum.Comments = datumItem.<Comments>.Value
                NewEngineeringDatum.OriginDescription = datumItem.<OriginDescription>.Value
                NewEngineeringDatum.Epoch = datumItem.<RealizationEpoch>.Value
                NewEngineeringDatum.Scope = datumItem.<Scope>.Value
                NewEngineeringDatum.Deprecated = datumItem.<Deprecated>.Value

                Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewEngineeringDatum.AddAlias(nameItem)
                Next

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewEngineeringDatum.Area.Name = datumItem.<AreaOfUse>.<Name>.Value
                NewEngineeringDatum.Area.Author = datumItem.<AreaOfUse>.<Author>.Value
                NewEngineeringDatum.Area.Code = datumItem.<AreaOfUse>.<Code>.Value

                List.Add(NewEngineeringDatum)
            Next
        End Sub

        Public Sub LoadFile()
            'Load the list from the selected list file.

            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<EngineeringDatumList>.<CreationDate>.Value
            LastEditDate = XDoc.<EngineeringDatumList>.<LastEditDate>.Value
            Description = XDoc.<EngineeringDatumList>.<Description>.Value

            Dim EngineeringDatums = From item In XDoc.<EngineeringDatumList>.<EngineeringDatum>

            List.Clear()
            For Each datumItem In EngineeringDatums
                Dim NewEngineeringDatum As New EngineeringDatum
                NewEngineeringDatum.Name = datumItem.<Name>.Value
                NewEngineeringDatum.Author = datumItem.<Author>.Value
                If datumItem.<Code>.Value = Nothing Then
                    NewEngineeringDatum.Code = 0
                Else
                    NewEngineeringDatum.Code = datumItem.<Code>.Value
                End If
                NewEngineeringDatum.Selected = datumItem.<Selected>.Value
                NewEngineeringDatum.Comments = datumItem.<Comments>.Value
                NewEngineeringDatum.OriginDescription = datumItem.<OriginDescription>.Value
                NewEngineeringDatum.Epoch = datumItem.<RealizationEpoch>.Value
                NewEngineeringDatum.Scope = datumItem.<Scope>.Value
                NewEngineeringDatum.Deprecated = datumItem.<Deprecated>.Value

                Dim aliasNames = From item In datumItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewEngineeringDatum.AddAlias(nameItem)
                Next

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewEngineeringDatum.Area.Name = datumItem.<AreaOfUse>.<Name>.Value
                NewEngineeringDatum.Area.Author = datumItem.<AreaOfUse>.<Author>.Value
                NewEngineeringDatum.Area.Code = datumItem.<AreaOfUse>.<Code>.Value

                List.Add(NewEngineeringDatum)
            Next

        End Sub

        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            'Load the Datum list from the EPSG database.

            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Datum table into dataset ds - Only read the engineering datum records.
            da = New OleDb.OleDbDataAdapter("Select * From [Datum] Where [DATUM_TYPE] = 'engineering'", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)

            ''Read the Ellipsoid table into dataset ds
            'da.SelectCommand = New OleDb.OleDbCommand("Select * From [Ellipsoid]", myConnection)
            'TableName = "Ellipsoid"
            'da.Fill(ds, TableName)

            ''Read the Prime Meridian table into dataset ds
            'da.SelectCommand = New OleDb.OleDbCommand("Select * From [Prime Meridian]", myConnection)
            'TableName = "Prime Meridian"
            'da.Fill(ds, TableName)

            ''Read the Unit of Measure table into ds
            'da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            'TableName = "Unit of Measure"
            'da.Fill(ds, TableName)
            ''Unit of Measure Table fields :
            ''UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            Dim UOMCode As Integer

            Dim NRows As Integer = ds.Tables("Datum").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Engineering Datum list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Engineering Datums from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the Datum table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                Dim NewDatum As New EngineeringDatum
                NewDatum.Name = ds.Tables("Datum").Rows(RowNo).Item("DATUM_NAME")
                NewDatum.Author = "EPSG"
                NewDatum.Code = ds.Tables("Datum").Rows(RowNo).Item("DATUM_CODE")
                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REMARKS")) Then
                    NewDatum.Comments = ""
                Else
                    NewDatum.Comments = ds.Tables("Datum").Rows(RowNo).Item("REMARKS")
                End If

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")) Then
                    NewDatum.OriginDescription = ""
                Else
                    NewDatum.OriginDescription = ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")
                End If

                If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")) Then
                    NewDatum.Epoch = ""
                Else
                    NewDatum.Epoch = ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")
                End If

                NewDatum.Scope = ds.Tables("Datum").Rows(RowNo).Item("DATUM_SCOPE")
                NewDatum.Deprecated = ds.Tables("Datum").Rows(RowNo).Item("DEPRECATED")

                'TO DO: CHECK OBJECT_TABLE_NAME AND OBJECT_CODE FOR EVERY LoadEpsgDbList METHOD
                'Add list of alias names --------------------------------------------------------------------------------------------
                expression = "[OBJECT_TABLE_NAME] = 'Datum' AND [OBJECT_CODE] = " & Str(NewDatum.Code)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewDatum.AddAlias(item.Item("ALIAS").ToString)
                Next
                '---------------------------------------------------------------------------------------------------------------------

                'Add Area of Use information -----------------------------------------------------------------------------------------
                NewDatum.Area.Code = ds.Tables("Datum").Rows(RowNo).Item("AREA_OF_USE_CODE")
                NewDatum.Area.Author = "EPSG"
                expression = "[AREA_CODE] = " & Str(NewDatum.Area.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewDatum.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                End If
                List.Add(NewDatum)
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Engineering Datums as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Engineering Datum List File-->
                       <EngineeringDatumList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <EngineeringDatum>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <OriginDescription><%= item.OriginDescription %></OriginDescription>
                                     <RealizationEpoch><%= item.Epoch %></RealizationEpoch>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </EngineeringDatum>
                           %>
                           <!---->
                       </EngineeringDatumList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'EngineeringDatumList ----------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class Geographic2DCRSList '--------------------------------------------------------------------------------------------------------------------------------------------------------

        Public List As New List(Of Geographic2DCRSSummary)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Geographic 2D CRS List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<Geographic2DCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<Geographic2DCRSList>.<LastEditDate>.Value
            Description = XDoc.<Geographic2DCRSList>.<Description>.Value

            Dim GeoCRSs = From item In XDoc.<Geographic2DCRSList>.<Geographic2DCRS>

            List.Clear()
            For Each crsItem In GeoCRSs
                Dim NewGeoCRS As New Geographic2DCRSSummary
                NewGeoCRS.Name = crsItem.<Name>.Value
                NewGeoCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewGeoCRS.Code = 0
                Else
                    NewGeoCRS.Code = crsItem.<Code>.Value
                End If
                NewGeoCRS.Selected = crsItem.<Selected>.Value
                NewGeoCRS.Comments = crsItem.<Comments>.Value
                NewGeoCRS.Scope = crsItem.<Scope>.Value
                NewGeoCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewGeoCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewGeoCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewGeoCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewGeoCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewGeoCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewGeoCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewGeoCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read SourceGeographicCRS information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.SourceGeographicCRS.Name = crsItem.<SourceGeographicCRS>.<Name>.Value
                NewGeoCRS.SourceGeographicCRS.Author = crsItem.<SourceGeographicCRS>.<Author>.Value
                NewGeoCRS.SourceGeographicCRS.Code = crsItem.<SourceGeographicCRS>.<Code>.Value
                NewGeoCRS.SourceGeographicCRS.Type = crsItem.<SourceGeographicCRS>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewGeoCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewGeoCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewGeoCRS)
            Next
        End Sub

        'Load the list from the selected list file.
        Public Sub LoadFile()
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<Geographic2DCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<Geographic2DCRSList>.<LastEditDate>.Value
            Description = XDoc.<Geographic2DCRSList>.<Description>.Value

            Dim GeoCRSs = From item In XDoc.<Geographic2DCRSList>.<Geographic2DCRS>

            List.Clear()
            For Each crsItem In GeoCRSs
                Dim NewGeoCRS As New Geographic2DCRSSummary
                NewGeoCRS.Name = crsItem.<Name>.Value
                NewGeoCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewGeoCRS.Code = 0
                Else
                    NewGeoCRS.Code = crsItem.<Code>.Value
                End If
                NewGeoCRS.Selected = crsItem.<Selected>.Value
                NewGeoCRS.Comments = crsItem.<Comments>.Value
                NewGeoCRS.Scope = crsItem.<Scope>.Value
                NewGeoCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewGeoCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewGeoCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewGeoCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewGeoCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewGeoCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewGeoCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewGeoCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read SourceGeographicCRS information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.SourceGeographicCRS.Name = crsItem.<SourceGeographicCRS>.<Name>.Value
                NewGeoCRS.SourceGeographicCRS.Author = crsItem.<SourceGeographicCRS>.<Author>.Value
                NewGeoCRS.SourceGeographicCRS.Code = crsItem.<SourceGeographicCRS>.<Code>.Value
                NewGeoCRS.SourceGeographicCRS.Type = crsItem.<SourceGeographicCRS>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewGeoCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewGeoCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewGeoCRS)
            Next

        End Sub

        'Load the CRS list from the EPSG database.
        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Coordinate Reference System table into dataset ds. All types are included. Base CRS data is sometimes required
            '(Only records where [COORD_REF_SYS_KIND ] = 'geographic 2D' will be processed.)
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System]", myConnection)
            TableName = "CoordRefSys"
            da.Fill(ds, TableName)
            'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
            'CMPD_HORIZCRS_CODE and CMPD_VERTCRS_CODE are only used for compound CRS types

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)
            'AREA_CODE AREA_NAME AREA_OF_USE AREA_SOUTH_BOUND_LAT AREA_NORTH_BOUND_LAT AREA_WEST_BOUND_LON AREA_EAST_BOUND_LON AREA_POLYGON_FILE_REF ISO_A2_CODE ISO_A3_CODE ISO_N_CODE REMARKS DEPRECATED

            'Read the Coordinate System table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate System]", myConnection)
            TableName = "CoordSys"
            da.Fill(ds, TableName)
            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED

            'Read the Datum table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Datum]", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)
            'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            'Dim CoordRefSysCode As Integer
            'Dim AreaOfUseCode As Integer
            'Dim CoordSysCode As Integer
            Dim TargetUOMCode As Integer
            'Dim DatumCode As Integer
            'Dim SourceGeogCrsCode As Integer
            Dim ProjectionConvCode As Integer

            Dim NRows As Integer = ds.Tables("CoordRefSys").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Geographic 2D CRS list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            'Dim EllipsoidParametersString As String
            RaiseEvent Message("Reading GGeographic 2D CRSs from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordRefSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                If ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_KIND") = "geographic 2D" Then 'Add this record to the list:
                    Dim NewCRS As New Geographic2DCRSSummary
                    NewCRS.Name = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_NAME")
                    NewCRS.Author = "EPSG"
                    NewCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_CODE")
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")) Then
                        NewCRS.Comments = ""
                    Else
                        NewCRS.Comments = ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")
                    End If

                    'If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")) Then
                    '    NewDatum.OriginDescription = ""
                    'Else
                    '    NewDatum.OriginDescription = ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")
                    'End If

                    'If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")) Then
                    '    NewDatum.Epoch = ""
                    'Else
                    '    NewDatum.Epoch = ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")
                    'End If

                    NewCRS.Scope = ds.Tables("CoordRefSys").Rows(RowNo).Item("CRS_SCOPE")
                    NewCRS.Deprecated = ds.Tables("CoordRefSys").Rows(RowNo).Item("DEPRECATED")

                    'Add list of alias names --------------------------------------------------------------------------------------------
                    'expression = "[OBJECT_TABLE_NAME] = 'Prime Meridian' AND [OBJECT_CODE] = " & Str(NewDatum.PrimeMeridian.Code)
                    expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.Code)
                    'Dim pmAliasResult = ds.Tables("Alias").Select(expression)
                    Dim result = ds.Tables("Alias").Select(expression)
                    For Each item In result
                        NewCRS.AddAlias(item.Item("ALIAS").ToString)
                    Next

                    'Add Coordinate System information -----------------------------------------------------------------------------------------
                    'Dim coordSystem As New XElement("CoordinateSystem")
                    NewCRS.CoordinateSystem.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_SYS_CODE")
                    expression = "[COORD_SYS_CODE] = " & Str(NewCRS.CoordinateSystem.Code)
                    Dim coordSysParameters = ds.Tables("CoordSys").Select(expression)
                    'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                    If coordSysParameters.Count > 0 Then
                        NewCRS.CoordinateSystem.Name = coordSysParameters(0).Item("COORD_SYS_NAME").ToString
                        NewCRS.CoordinateSystem.Author = "EPSG"
                        NewCRS.CoordinateSystem.Type = coordSysParameters(0).Item("COORD_SYS_TYPE").ToString
                    End If

                    'Add Datum information -----------------------------------------------------------------------------------------
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")) Then
                        NewCRS.Datum.Code = 0
                    Else
                        NewCRS.Datum.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")
                    End If

                    expression = "[DATUM_CODE] = " & Str(NewCRS.Datum.Code)
                    Dim datumParameters = ds.Tables("Datum").Select(expression)
                    'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED
                    If datumParameters.Count > 0 Then
                        NewCRS.Datum.Name = datumParameters(0).Item("DATUM_NAME")
                        NewCRS.Datum.Author = "EPSG"
                        NewCRS.Datum.Type = datumParameters(0).Item("DATUM_TYPE")
                    End If

                    'Add Source Geographic Coordinate System information -----------------------------------------------------------------------------------------
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("SOURCE_GEOGCRS_CODE")) Then
                        NewCRS.SourceGeographicCRS.Code = 0
                    Else
                        NewCRS.SourceGeographicCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("SOURCE_GEOGCRS_CODE")
                    End If

                    expression = "[SOURCE_GEOGCRS_CODE] = " & Str(NewCRS.SourceGeographicCRS.Code)
                    Dim coordRefSysParameters = ds.Tables("CoordRefSys").Select(expression)
                    'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
                    If coordRefSysParameters.Count > 0 Then
                        NewCRS.SourceGeographicCRS.Name = coordRefSysParameters(0).Item("COORD_REF_SYS_NAME")
                        NewCRS.SourceGeographicCRS.Author = "EPSG"
                        NewCRS.SourceGeographicCRS.Type = coordRefSysParameters(0).Item("COORD_REF_SYS_KIND")
                    End If

                    'Add Area of Use information -----------------------------------------------------------------------------------------
                    NewCRS.Area.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("AREA_OF_USE_CODE")
                    expression = "[AREA_CODE] = " & Str(NewCRS.Area.Code)
                    Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                    If areaOfUseParameters.Count > 0 Then
                        NewCRS.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                        NewCRS.Area.Author = "EPSG"
                    End If
                    List.Add(NewCRS)
                End If
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Geographic 2D CRS as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Geographic 2D Coordinate Reference System List File-->
                       <Geographic2DCRSList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <Geographic2DCRS>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <Datum>
                                         <Name><%= item.Datum.Name %></Name>
                                         <Author><%= item.Datum.Author %></Author>
                                         <Code><%= item.Datum.Code %></Code>
                                         <Type><%= item.Datum.Type %></Type>
                                     </Datum>
                                     <CoordinateSystem>
                                         <Name><%= item.CoordinateSystem.Name %></Name>
                                         <Author><%= item.CoordinateSystem.Author %></Author>
                                         <Code><%= item.CoordinateSystem.Code %></Code>
                                         <Type><%= item.CoordinateSystem.Type %></Type>
                                     </CoordinateSystem>
                                     <SourceGeographicCRS>
                                         <Name><%= item.SourceGeographicCRS.Name %></Name>
                                         <Author><%= item.SourceGeographicCRS.Author %></Author>
                                         <Code><%= item.SourceGeographicCRS.Code %></Code>
                                         <Type><%= item.SourceGeographicCRS.Type %></Type>
                                     </SourceGeographicCRS>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </Geographic2DCRS>
                           %>
                           <!---->
                       </Geographic2DCRSList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'Geographic2DCRSList -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class Geographic3DCRSList '--------------------------------------------------------------------------------------------------------------------------------------------------------
        Public List As New List(Of Geographic3DCRSSummary)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Geographic 2D CRS List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<Geographic3DCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<Geographic3DCRSList>.<LastEditDate>.Value
            Description = XDoc.<Geographic3DCRSList>.<Description>.Value

            Dim GeoCRSs = From item In XDoc.<Geographic3DCRSList>.<Geographic3DCRS>

            List.Clear()
            For Each crsItem In GeoCRSs
                Dim NewGeoCRS As New Geographic3DCRSSummary
                NewGeoCRS.Name = crsItem.<Name>.Value
                NewGeoCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewGeoCRS.Code = 0
                Else
                    NewGeoCRS.Code = crsItem.<Code>.Value
                End If
                NewGeoCRS.Selected = crsItem.<Selected>.Value
                NewGeoCRS.Comments = crsItem.<Comments>.Value
                NewGeoCRS.Scope = crsItem.<Scope>.Value
                NewGeoCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewGeoCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewGeoCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewGeoCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewGeoCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewGeoCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewGeoCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewGeoCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read SourceGeographicCRS information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.SourceGeographicCRS.Name = crsItem.<SourceGeographicCRS>.<Name>.Value
                NewGeoCRS.SourceGeographicCRS.Author = crsItem.<SourceGeographicCRS>.<Author>.Value
                NewGeoCRS.SourceGeographicCRS.Code = crsItem.<SourceGeographicCRS>.<Code>.Value
                NewGeoCRS.SourceGeographicCRS.Type = crsItem.<SourceGeographicCRS>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewGeoCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewGeoCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewGeoCRS)
            Next
        End Sub

        'Load the list from the selected list file.
        Public Sub LoadFile()
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<Geographic3DCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<Geographic3DCRSList>.<LastEditDate>.Value
            Description = XDoc.<Geographic3DCRSList>.<Description>.Value

            Dim GeoCRSs = From item In XDoc.<Geographic3DCRSList>.<Geographic3DCRS>

            List.Clear()
            For Each crsItem In GeoCRSs
                Dim NewGeoCRS As New Geographic3DCRSSummary
                NewGeoCRS.Name = crsItem.<Name>.Value
                NewGeoCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewGeoCRS.Code = 0
                Else
                    NewGeoCRS.Code = crsItem.<Code>.Value
                End If
                NewGeoCRS.Selected = crsItem.<Selected>.Value
                NewGeoCRS.Comments = crsItem.<Comments>.Value
                NewGeoCRS.Scope = crsItem.<Scope>.Value
                NewGeoCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewGeoCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewGeoCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewGeoCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewGeoCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewGeoCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewGeoCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewGeoCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read SourceGeographicCRS information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.SourceGeographicCRS.Name = crsItem.<SourceGeographicCRS>.<Name>.Value
                NewGeoCRS.SourceGeographicCRS.Author = crsItem.<SourceGeographicCRS>.<Author>.Value
                NewGeoCRS.SourceGeographicCRS.Code = crsItem.<SourceGeographicCRS>.<Code>.Value
                NewGeoCRS.SourceGeographicCRS.Type = crsItem.<SourceGeographicCRS>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewGeoCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewGeoCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewGeoCRS)
            Next

        End Sub

        'Load the CRS list from the EPSG database.
        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Coordinate Reference System table into dataset ds. All types are included. Base CRS data is sometimes required
            '(Only records where [COORD_REF_SYS_KIND ] = 'geographic 3D' will be processed.)
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System]", myConnection)
            TableName = "CoordRefSys"
            da.Fill(ds, TableName)
            'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
            'CMPD_HORIZCRS_CODE and CMPD_VERTCRS_CODE are only used for compound CRS types

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)
            'AREA_CODE AREA_NAME AREA_OF_USE AREA_SOUTH_BOUND_LAT AREA_NORTH_BOUND_LAT AREA_WEST_BOUND_LON AREA_EAST_BOUND_LON AREA_POLYGON_FILE_REF ISO_A2_CODE ISO_A3_CODE ISO_N_CODE REMARKS DEPRECATED

            'Read the Coordinate System table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate System]", myConnection)
            TableName = "CoordSys"
            da.Fill(ds, TableName)
            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED

            'Read the Datum table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Datum]", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)
            'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            'Dim CoordRefSysCode As Integer
            'Dim AreaOfUseCode As Integer
            'Dim CoordSysCode As Integer
            Dim TargetUOMCode As Integer
            'Dim DatumCode As Integer
            'Dim SourceGeogCrsCode As Integer
            Dim ProjectionConvCode As Integer

            Dim NRows As Integer = ds.Tables("CoordRefSys").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Geographic 3D CRS list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            'Dim EllipsoidParametersString As String
            RaiseEvent Message("Reading GGeographic 3D CRSs from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordRefSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                If ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_KIND") = "geographic 3D" Then 'Add this record to the list:
                    Dim NewCRS As New Geographic3DCRSSummary
                    NewCRS.Name = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_NAME")
                    NewCRS.Author = "EPSG"
                    NewCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_CODE")
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")) Then
                        NewCRS.Comments = ""
                    Else
                        NewCRS.Comments = ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")
                    End If

                    'If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")) Then
                    '    NewDatum.OriginDescription = ""
                    'Else
                    '    NewDatum.OriginDescription = ds.Tables("Datum").Rows(RowNo).Item("ORIGIN_DESCRIPTION")
                    'End If

                    'If IsDBNull(ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")) Then
                    '    NewDatum.Epoch = ""
                    'Else
                    '    NewDatum.Epoch = ds.Tables("Datum").Rows(RowNo).Item("REALIZATION_EPOCH")
                    'End If

                    NewCRS.Scope = ds.Tables("CoordRefSys").Rows(RowNo).Item("CRS_SCOPE")
                    NewCRS.Deprecated = ds.Tables("CoordRefSys").Rows(RowNo).Item("DEPRECATED")

                    'Add list of alias names --------------------------------------------------------------------------------------------
                    'expression = "[OBJECT_TABLE_NAME] = 'Prime Meridian' AND [OBJECT_CODE] = " & Str(NewDatum.PrimeMeridian.Code)
                    expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.Code)
                    'Dim pmAliasResult = ds.Tables("Alias").Select(expression)
                    Dim result = ds.Tables("Alias").Select(expression)
                    For Each item In result
                        NewCRS.AddAlias(item.Item("ALIAS").ToString)
                    Next

                    'Add Coordinate System information -----------------------------------------------------------------------------------------
                    'Dim coordSystem As New XElement("CoordinateSystem")
                    NewCRS.CoordinateSystem.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_SYS_CODE")
                    expression = "[COORD_SYS_CODE] = " & Str(NewCRS.CoordinateSystem.Code)
                    Dim coordSysParameters = ds.Tables("CoordSys").Select(expression)
                    'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                    If coordSysParameters.Count > 0 Then
                        NewCRS.CoordinateSystem.Name = coordSysParameters(0).Item("COORD_SYS_NAME").ToString
                        NewCRS.CoordinateSystem.Author = "EPSG"
                        NewCRS.CoordinateSystem.Type = coordSysParameters(0).Item("COORD_SYS_TYPE").ToString
                    End If

                    'Add Datum information -----------------------------------------------------------------------------------------
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")) Then
                        NewCRS.Datum.Code = 0
                    Else
                        NewCRS.Datum.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")
                    End If

                    expression = "[DATUM_CODE] = " & Str(NewCRS.Datum.Code)
                    Dim datumParameters = ds.Tables("Datum").Select(expression)
                    'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED
                    If datumParameters.Count > 0 Then
                        NewCRS.Datum.Name = datumParameters(0).Item("DATUM_NAME")
                        NewCRS.Datum.Author = "EPSG"
                        NewCRS.Datum.Type = datumParameters(0).Item("DATUM_TYPE")
                    End If

                    'Add Source Geographic Coordinate System information -----------------------------------------------------------------------------------------
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("SOURCE_GEOGCRS_CODE")) Then
                        NewCRS.SourceGeographicCRS.Code = 0
                    Else
                        NewCRS.SourceGeographicCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("SOURCE_GEOGCRS_CODE")
                    End If

                    expression = "[SOURCE_GEOGCRS_CODE] = " & Str(NewCRS.SourceGeographicCRS.Code)
                    Dim coordRefSysParameters = ds.Tables("CoordRefSys").Select(expression)
                    'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
                    If coordRefSysParameters.Count > 0 Then
                        NewCRS.SourceGeographicCRS.Name = coordRefSysParameters(0).Item("COORD_REF_SYS_NAME")
                        NewCRS.SourceGeographicCRS.Author = "EPSG"
                        NewCRS.SourceGeographicCRS.Type = coordRefSysParameters(0).Item("COORD_REF_SYS_KIND")
                    End If

                    'Add Area of Use information -----------------------------------------------------------------------------------------
                    NewCRS.Area.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("AREA_OF_USE_CODE")
                    expression = "[AREA_CODE] = " & Str(NewCRS.Area.Code)
                    Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                    If areaOfUseParameters.Count > 0 Then
                        NewCRS.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                        NewCRS.Area.Author = "EPSG"
                    End If
                    List.Add(NewCRS)
                End If
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Geographic 2D CRS as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Geographic 2D Coordinate Reference System List File-->
                       <Geographic3DCRSList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <Geographic3DCRS>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <Datum>
                                         <Name><%= item.Datum.Name %></Name>
                                         <Author><%= item.Datum.Author %></Author>
                                         <Code><%= item.Datum.Code %></Code>
                                         <Type><%= item.Datum.Type %></Type>
                                     </Datum>
                                     <CoordinateSystem>
                                         <Name><%= item.CoordinateSystem.Name %></Name>
                                         <Author><%= item.CoordinateSystem.Author %></Author>
                                         <Code><%= item.CoordinateSystem.Code %></Code>
                                         <Type><%= item.CoordinateSystem.Type %></Type>
                                     </CoordinateSystem>
                                     <SourceGeographicCRS>
                                         <Name><%= item.SourceGeographicCRS.Name %></Name>
                                         <Author><%= item.SourceGeographicCRS.Author %></Author>
                                         <Code><%= item.SourceGeographicCRS.Code %></Code>
                                         <Type><%= item.SourceGeographicCRS.Type %></Type>
                                     </SourceGeographicCRS>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </Geographic3DCRS>
                           %>
                           <!---->
                       </Geographic3DCRSList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------
    End Class 'Geographic3DCRSList -----------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class GeocentricCRSList '----------------------------------------------------------------------------------------------------------------------------------------------------------
        Public List As New List(Of GeocentricCRSSummary)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property
#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Geocentric CRS List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<GeocentricCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<GeocentricCRSList>.<LastEditDate>.Value
            Description = XDoc.<GeocentricCRSList>.<Description>.Value

            Dim GeoCRSs = From item In XDoc.<GeocentricCRSList>.<GeocentricCRS>

            List.Clear()
            For Each crsItem In GeoCRSs
                Dim NewGeoCRS As New GeocentricCRSSummary
                NewGeoCRS.Name = crsItem.<Name>.Value
                NewGeoCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewGeoCRS.Code = 0
                Else
                    NewGeoCRS.Code = crsItem.<Code>.Value
                End If
                NewGeoCRS.Selected = crsItem.<Selected>.Value
                NewGeoCRS.Comments = crsItem.<Comments>.Value
                NewGeoCRS.Scope = crsItem.<Scope>.Value
                NewGeoCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewGeoCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewGeoCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewGeoCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewGeoCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewGeoCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewGeoCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewGeoCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewGeoCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewGeoCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewGeoCRS)
            Next
        End Sub

        'Load the list from the selected list file.
        Public Sub LoadFile()
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<GeocentricCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<GeocentricCRSList>.<LastEditDate>.Value
            Description = XDoc.<GeocentricCRSList>.<Description>.Value

            Dim GeoCRSs = From item In XDoc.<GeocentricCRSList>.<GeocentricCRS>

            List.Clear()
            For Each crsItem In GeoCRSs
                Dim NewGeoCRS As New GeocentricCRSSummary
                NewGeoCRS.Name = crsItem.<Name>.Value
                NewGeoCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewGeoCRS.Code = 0
                Else
                    NewGeoCRS.Code = crsItem.<Code>.Value
                End If
                NewGeoCRS.Selected = crsItem.<Selected>.Value
                NewGeoCRS.Comments = crsItem.<Comments>.Value
                NewGeoCRS.Scope = crsItem.<Scope>.Value
                NewGeoCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewGeoCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewGeoCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewGeoCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewGeoCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewGeoCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewGeoCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewGeoCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewGeoCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewGeoCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewGeoCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewGeoCRS)
            Next
        End Sub

        'Load the CRS list from the EPSG database.
        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Only records where COORD_REF_SYS_KIND = 'geocentric' will be processed.)
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System] WHERE COORD_REF_SYS_KIND = 'geocentric'", myConnection)
            TableName = "CoordRefSys"
            da.Fill(ds, TableName)
            'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
            'CMPD_HORIZCRS_CODE and CMPD_VERTCRS_CODE are only used for compound CRS types

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)
            'AREA_CODE AREA_NAME AREA_OF_USE AREA_SOUTH_BOUND_LAT AREA_NORTH_BOUND_LAT AREA_WEST_BOUND_LON AREA_EAST_BOUND_LON AREA_POLYGON_FILE_REF ISO_A2_CODE ISO_A3_CODE ISO_N_CODE REMARKS DEPRECATED

            'Read the Coordinate System table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate System]", myConnection)
            TableName = "CoordSys"
            da.Fill(ds, TableName)
            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED

            'Read the Datum table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Datum]", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)
            'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            'Dim AreaOfUseCode As Integer
            'Dim CoordSysCode As Integer
            Dim TargetUOMCode As Integer
            'Dim DatumCode As Integer
            'Dim ProjectionConvCode As Integer

            Dim NRows As Integer = ds.Tables("CoordRefSys").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Geocentric CRS list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Geocentric CRSs from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordRefSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                'If ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_KIND") = "geographic 2D" Then 'Add this record to the list:
                Dim NewCRS As New GeocentricCRSSummary
                NewCRS.Name = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_NAME")
                NewCRS.Author = "EPSG"
                NewCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_CODE")
                If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")) Then
                    NewCRS.Comments = ""
                Else
                    NewCRS.Comments = ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")
                End If

                NewCRS.Scope = ds.Tables("CoordRefSys").Rows(RowNo).Item("CRS_SCOPE")
                NewCRS.Deprecated = ds.Tables("CoordRefSys").Rows(RowNo).Item("DEPRECATED")

                'Add list of alias names --------------------------------------------------------------------------------------------
                'expression = "[OBJECT_TABLE_NAME] = 'Prime Meridian' AND [OBJECT_CODE] = " & Str(NewDatum.PrimeMeridian.Code)
                expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.Code)
                'Dim pmAliasResult = ds.Tables("Alias").Select(expression)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewCRS.AddAlias(item.Item("ALIAS").ToString)
                Next

                'Add Coordinate System information -----------------------------------------------------------------------------------------
                'Dim coordSystem As New XElement("CoordinateSystem")
                NewCRS.CoordinateSystem.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_SYS_CODE")
                expression = "[COORD_SYS_CODE] = " & Str(NewCRS.CoordinateSystem.Code)
                Dim coordSysParameters = ds.Tables("CoordSys").Select(expression)
                'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                If coordSysParameters.Count > 0 Then
                    NewCRS.CoordinateSystem.Name = coordSysParameters(0).Item("COORD_SYS_NAME").ToString
                    NewCRS.CoordinateSystem.Author = "EPSG"
                    NewCRS.CoordinateSystem.Type = coordSysParameters(0).Item("COORD_SYS_TYPE").ToString
                End If

                'Add Datum information -----------------------------------------------------------------------------------------
                If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")) Then
                    NewCRS.Datum.Code = 0
                Else
                    NewCRS.Datum.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")
                End If

                expression = "[DATUM_CODE] = " & Str(NewCRS.Datum.Code)
                Dim datumParameters = ds.Tables("Datum").Select(expression)
                'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED
                If datumParameters.Count > 0 Then
                    NewCRS.Datum.Name = datumParameters(0).Item("DATUM_NAME")
                    NewCRS.Datum.Author = "EPSG"
                    NewCRS.Datum.Type = datumParameters(0).Item("DATUM_TYPE")
                End If

                ''Add Source Geographic Coordinate System information -----------------------------------------------------------------------------------------
                'If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("SOURCE_GEOGCRS_CODE")) Then
                '    NewCRS.SourceGeographicCRS.Code = 0
                'Else
                '    NewCRS.SourceGeographicCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("SOURCE_GEOGCRS_CODE")
                'End If

                'expression = "[SOURCE_GEOGCRS_CODE] = " & Str(NewCRS.SourceGeographicCRS.Code)
                'Dim coordRefSysParameters = ds.Tables("CoordRefSys").Select(expression)
                ''COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
                'If coordRefSysParameters.Count > 0 Then
                '    NewCRS.SourceGeographicCRS.Name = coordRefSysParameters(0).Item("COORD_REF_SYS_NAME")
                '    NewCRS.SourceGeographicCRS.Author = "EPSG"
                '    NewCRS.SourceGeographicCRS.Type = coordRefSysParameters(0).Item("COORD_REF_SYS_KIND")
                'End If

                'Add Area of Use information -----------------------------------------------------------------------------------------
                NewCRS.Area.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("AREA_OF_USE_CODE")
                expression = "[AREA_CODE] = " & Str(NewCRS.Area.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewCRS.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                    NewCRS.Area.Author = "EPSG"
                End If
                List.Add(NewCRS)
                'End If
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Geocentric CRS as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Geocentric Coordinate Reference System List File-->
                       <GeocentricCRSList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <GeocentricCRS>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <Datum>
                                         <Name><%= item.Datum.Name %></Name>
                                         <Author><%= item.Datum.Author %></Author>
                                         <Code><%= item.Datum.Code %></Code>
                                         <Type><%= item.Datum.Type %></Type>
                                     </Datum>
                                     <CoordinateSystem>
                                         <Name><%= item.CoordinateSystem.Name %></Name>
                                         <Author><%= item.CoordinateSystem.Author %></Author>
                                         <Code><%= item.CoordinateSystem.Code %></Code>
                                         <Type><%= item.CoordinateSystem.Type %></Type>
                                     </CoordinateSystem>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </GeocentricCRS>
                           %>
                           <!---->
                       </GeocentricCRSList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'GeocentricCRSList -------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class VerticalCRSList '------------------------------------------------------------------------------------------------------------------------------------------------------------
        Public List As New List(Of VerticalCRSSummary)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property
#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Vertical CRS List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<VerticalCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<VerticalCRSList>.<LastEditDate>.Value
            Description = XDoc.<VerticalCRSList>.<Description>.Value

            Dim VertCRSs = From item In XDoc.<VerticalCRSList>.<VerticalCRS>

            List.Clear()
            For Each crsItem In VertCRSs
                Dim NewVertCRS As New VerticalCRSSummary
                NewVertCRS.Name = crsItem.<Name>.Value
                NewVertCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewVertCRS.Code = 0
                Else
                    NewVertCRS.Code = crsItem.<Code>.Value
                End If
                NewVertCRS.Selected = crsItem.<Selected>.Value
                NewVertCRS.Comments = crsItem.<Comments>.Value
                NewVertCRS.Scope = crsItem.<Scope>.Value
                NewVertCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewVertCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewVertCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewVertCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewVertCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewVertCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewVertCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewVertCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewVertCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewVertCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewVertCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewVertCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewVertCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewVertCRS)
            Next
        End Sub

        'Load the list from the selected list file.
        Public Sub LoadFile()
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<VerticalCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<VerticalCRSList>.<LastEditDate>.Value
            Description = XDoc.<VerticalCRSList>.<Description>.Value

            Dim VertCRSs = From item In XDoc.<VerticalCRSList>.<VerticalCRS>

            List.Clear()
            For Each crsItem In VertCRSs
                Dim NewVertCRS As New VerticalCRSSummary
                NewVertCRS.Name = crsItem.<Name>.Value
                NewVertCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewVertCRS.Code = 0
                Else
                    NewVertCRS.Code = crsItem.<Code>.Value
                End If
                NewVertCRS.Selected = crsItem.<Selected>.Value
                NewVertCRS.Comments = crsItem.<Comments>.Value
                NewVertCRS.Scope = crsItem.<Scope>.Value
                NewVertCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewVertCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewVertCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewVertCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewVertCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewVertCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewVertCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewVertCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewVertCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewVertCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewVertCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewVertCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewVertCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewVertCRS)
            Next
        End Sub

        'Load the CRS list from the EPSG database.
        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Only records where COORD_REF_SYS_KIND = 'vertical' will be processed.)
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System] WHERE COORD_REF_SYS_KIND = 'vertical'", myConnection)
            TableName = "CoordRefSys"
            da.Fill(ds, TableName)
            'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
            'CMPD_HORIZCRS_CODE and CMPD_VERTCRS_CODE are only used for compound CRS types

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)
            'AREA_CODE AREA_NAME AREA_OF_USE AREA_SOUTH_BOUND_LAT AREA_NORTH_BOUND_LAT AREA_WEST_BOUND_LON AREA_EAST_BOUND_LON AREA_POLYGON_FILE_REF ISO_A2_CODE ISO_A3_CODE ISO_N_CODE REMARKS DEPRECATED

            'Read the Coordinate System table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate System]", myConnection)
            TableName = "CoordSys"
            da.Fill(ds, TableName)
            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED

            'Read the Datum table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Datum]", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)
            'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            'Dim AreaOfUseCode As Integer
            'Dim CoordSysCode As Integer
            Dim TargetUOMCode As Integer
            'Dim DatumCode As Integer
            'Dim ProjectionConvCode As Integer

            Dim NRows As Integer = ds.Tables("CoordRefSys").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Vertical CRS list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Vertical CRSs from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordRefSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                'If ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_KIND") = "geographic 2D" Then 'Add this record to the list:
                Dim NewCRS As New VerticalCRSSummary
                NewCRS.Name = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_NAME")
                NewCRS.Author = "EPSG"
                NewCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_CODE")
                If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")) Then
                    NewCRS.Comments = ""
                Else
                    NewCRS.Comments = ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")
                End If

                NewCRS.Scope = ds.Tables("CoordRefSys").Rows(RowNo).Item("CRS_SCOPE")
                NewCRS.Deprecated = ds.Tables("CoordRefSys").Rows(RowNo).Item("DEPRECATED")

                'Add list of alias names --------------------------------------------------------------------------------------------
                'expression = "[OBJECT_TABLE_NAME] = 'Prime Meridian' AND [OBJECT_CODE] = " & Str(NewDatum.PrimeMeridian.Code)
                expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.Code)
                'Dim pmAliasResult = ds.Tables("Alias").Select(expression)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewCRS.AddAlias(item.Item("ALIAS").ToString)
                Next

                'Add Coordinate System information -----------------------------------------------------------------------------------------
                'Dim coordSystem As New XElement("CoordinateSystem")
                NewCRS.CoordinateSystem.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_SYS_CODE")
                expression = "[COORD_SYS_CODE] = " & Str(NewCRS.CoordinateSystem.Code)
                Dim coordSysParameters = ds.Tables("CoordSys").Select(expression)
                'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                If coordSysParameters.Count > 0 Then
                    NewCRS.CoordinateSystem.Name = coordSysParameters(0).Item("COORD_SYS_NAME").ToString
                    NewCRS.CoordinateSystem.Author = "EPSG"
                    NewCRS.CoordinateSystem.Type = coordSysParameters(0).Item("COORD_SYS_TYPE").ToString
                End If

                'Add Datum information -----------------------------------------------------------------------------------------
                If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")) Then
                    NewCRS.Datum.Code = 0
                Else
                    NewCRS.Datum.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")
                End If

                expression = "[DATUM_CODE] = " & Str(NewCRS.Datum.Code)
                Dim datumParameters = ds.Tables("Datum").Select(expression)
                'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED
                If datumParameters.Count > 0 Then
                    NewCRS.Datum.Name = datumParameters(0).Item("DATUM_NAME")
                    NewCRS.Datum.Author = "EPSG"
                    NewCRS.Datum.Type = datumParameters(0).Item("DATUM_TYPE")
                End If

                'Add Area of Use information -----------------------------------------------------------------------------------------
                NewCRS.Area.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("AREA_OF_USE_CODE")
                expression = "[AREA_CODE] = " & Str(NewCRS.Area.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewCRS.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                    NewCRS.Area.Author = "EPSG"
                End If
                List.Add(NewCRS)
                'End If
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Vertical CRS as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Vertical Coordinate Reference System List File-->
                       <VerticalCRSList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <VerticalCRS>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <Datum>
                                         <Name><%= item.Datum.Name %></Name>
                                         <Author><%= item.Datum.Author %></Author>
                                         <Code><%= item.Datum.Code %></Code>
                                         <Type><%= item.Datum.Type %></Type>
                                     </Datum>
                                     <CoordinateSystem>
                                         <Name><%= item.CoordinateSystem.Name %></Name>
                                         <Author><%= item.CoordinateSystem.Author %></Author>
                                         <Code><%= item.CoordinateSystem.Code %></Code>
                                         <Type><%= item.CoordinateSystem.Type %></Type>
                                     </CoordinateSystem>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </VerticalCRS>
                           %>
                           <!---->
                       </VerticalCRSList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'VerticalCRSList ---------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class CompoundCRSList '------------------------------------------------------------------------------------------------------------------------------------------------------------
        Public List As New List(Of CompoundCRSSummary)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Compound CRS List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<CompoundCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<CompoundCRSList>.<LastEditDate>.Value
            Description = XDoc.<CompoundCRSList>.<Description>.Value

            Dim CompoundCRSs = From item In XDoc.<CompoundCRSList>.<CompoundCRS>

            List.Clear()
            For Each crsItem In CompoundCRSs
                Dim NewCompoundCRS As New CompoundCRSSummary
                NewCompoundCRS.Name = crsItem.<Name>.Value
                NewCompoundCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewCompoundCRS.Code = 0
                Else
                    NewCompoundCRS.Code = crsItem.<Code>.Value
                End If
                NewCompoundCRS.Selected = crsItem.<Selected>.Value
                NewCompoundCRS.Comments = crsItem.<Comments>.Value
                NewCompoundCRS.Scope = crsItem.<Scope>.Value
                NewCompoundCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewCompoundCRS.AddAlias(nameItem)
                Next


                'Read Horizontal CRS information: ----------------------------------------------------------------------------------------------
                NewCompoundCRS.HorizontalCRS.Name = crsItem.<HorizontalCRS>.<Name>.Value
                NewCompoundCRS.HorizontalCRS.Author = crsItem.<HorizontalCRS>.<Author>.Value
                If crsItem.<HorizontalCRS>.<Code>.Value = Nothing Then
                    NewCompoundCRS.HorizontalCRS.Code = 0
                Else
                    NewCompoundCRS.HorizontalCRS.Code = crsItem.<HorizontalCRS>.<Code>.Value
                End If
                NewCompoundCRS.HorizontalCRS.Comments = crsItem.<HorizontalCRS>.<Comments>.Value
                NewCompoundCRS.HorizontalCRS.Scope = crsItem.<HorizontalCRS>.<Scope>.Value
                NewCompoundCRS.HorizontalCRS.Deprecated = crsItem.<HorizontalCRS>.<Deprecated>.Value
                Dim HorizontalCRSAliasNames = From item In crsItem.<HorizontalCRS>.<AliasNames>.<AliasName>

                For Each nameItem In HorizontalCRSAliasNames
                    NewCompoundCRS.HorizontalCRS.AddAlias(nameItem)
                Next

                Select Case crsItem.<HorizontalCRS>.<Type>.Value
                    Case "Compound"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Compound
                    Case "Engineering"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Engineering
                    Case "Geocentric"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Geocentric
                    Case "Geographic2D"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Geographic2D
                    Case "Geographic3D"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Geographic3D
                    Case "Projected"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Projected
                    Case "Vertical"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Vertical
                    Case Else
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Unknown
                End Select

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewCompoundCRS.Area.Name = crsItem.<Area>.<Name>.Value
                NewCompoundCRS.Area.Author = crsItem.<Area>.<Author>.Value
                NewCompoundCRS.Area.Code = crsItem.<Area>.<Code>.Value

                NewCompoundCRS.HorizontalCRS.CoordinateSystem.Name = crsItem.<HorizontalCRS>.<CoordinateSystem>.<Name>.Value
                NewCompoundCRS.HorizontalCRS.CoordinateSystem.Author = crsItem.<HorizontalCRS>.<CoordinateSystem>.<Author>.Value
                NewCompoundCRS.HorizontalCRS.CoordinateSystem.Code = crsItem.<HorizontalCRS>.<CoordinateSystem>.<Code>.Value
                NewCompoundCRS.HorizontalCRS.CoordinateSystem.Type = crsItem.<HorizontalCRS>.<CoordinateSystem>.<Type>.Value

                NewCompoundCRS.HorizontalCRS.Area.Name = crsItem.<HorizontalCRS>.<Area>.<Name>.Value
                NewCompoundCRS.HorizontalCRS.Area.Author = crsItem.<HorizontalCRS>.<Area>.<Author>.Value
                NewCompoundCRS.HorizontalCRS.Area.Code = crsItem.<HorizontalCRS>.<Area>.<Code>.Value


                'Read Vertical CRS information: ------------------------------------------------------------------------------------------------
                NewCompoundCRS.VerticalCRS.Name = crsItem.<VerticalCRS>.<Name>.Value
                NewCompoundCRS.VerticalCRS.Author = crsItem.<VerticalCRS>.<Author>.Value
                If crsItem.<VerticalCRS>.<Code>.Value = Nothing Then
                    NewCompoundCRS.VerticalCRS.Code = 0
                Else
                    NewCompoundCRS.VerticalCRS.Code = crsItem.<VerticalCRS>.<Code>.Value
                End If
                NewCompoundCRS.VerticalCRS.Comments = crsItem.<VerticalCRS>.<Comments>.Value
                NewCompoundCRS.VerticalCRS.Scope = crsItem.<VerticalCRS>.<Scope>.Value
                NewCompoundCRS.VerticalCRS.Deprecated = crsItem.<VerticalCRS>.<Deprecated>.Value
                Dim VerticalCRSAliasNames = From item In crsItem.<VerticalCRS>.<AliasNames>.<AliasName>

                For Each nameItem In VerticalCRSAliasNames
                    NewCompoundCRS.VerticalCRS.AddAlias(nameItem)
                Next

                NewCompoundCRS.VerticalCRS.CoordinateSystem.Name = crsItem.<VerticalCRS>.<CoordinateSystem>.<Name>.Value
                NewCompoundCRS.VerticalCRS.CoordinateSystem.Author = crsItem.<VerticalCRS>.<CoordinateSystem>.<Author>.Value
                NewCompoundCRS.VerticalCRS.CoordinateSystem.Code = crsItem.<VerticalCRS>.<CoordinateSystem>.<Code>.Value
                NewCompoundCRS.VerticalCRS.CoordinateSystem.Type = crsItem.<VerticalCRS>.<CoordinateSystem>.<Type>.Value

                NewCompoundCRS.VerticalCRS.Datum.Name = crsItem.<VerticalCRS>.<Datum>.<Name>.Value
                NewCompoundCRS.VerticalCRS.Datum.Author = crsItem.<VerticalCRS>.<Datum>.<Author>.Value
                NewCompoundCRS.VerticalCRS.Datum.Code = crsItem.<VerticalCRS>.<Datum>.<Code>.Value

                NewCompoundCRS.VerticalCRS.Area.Name = crsItem.<VerticalCRS>.<Area>.<Name>.Value
                NewCompoundCRS.VerticalCRS.Area.Author = crsItem.<VerticalCRS>.<Area>.<Author>.Value
                NewCompoundCRS.VerticalCRS.Area.Code = crsItem.<VerticalCRS>.<Area>.<Code>.Value


               

                List.Add(NewCompoundCRS)
            Next
        End Sub

        'Load the list from the selected list file.
        Public Sub LoadFile()
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<CompoundCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<CompoundCRSList>.<LastEditDate>.Value
            Description = XDoc.<CompoundCRSList>.<Description>.Value

            Dim CompoundCRSs = From item In XDoc.<CompoundCRSList>.<CompoundCRS>

            List.Clear()
            For Each crsItem In CompoundCRSs
                Dim NewCompoundCRS As New CompoundCRSSummary
                NewCompoundCRS.Name = crsItem.<Name>.Value
                NewCompoundCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewCompoundCRS.Code = 0
                Else
                    NewCompoundCRS.Code = crsItem.<Code>.Value
                End If
                NewCompoundCRS.Selected = crsItem.<Selected>.Value
                NewCompoundCRS.Comments = crsItem.<Comments>.Value
                NewCompoundCRS.Scope = crsItem.<Scope>.Value
                NewCompoundCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewCompoundCRS.AddAlias(nameItem)
                Next


                'Read Horizontal CRS information: ----------------------------------------------------------------------------------------------
                NewCompoundCRS.HorizontalCRS.Name = crsItem.<HorizontalCRS>.<Name>.Value
                NewCompoundCRS.HorizontalCRS.Author = crsItem.<HorizontalCRS>.<Author>.Value
                If crsItem.<HorizontalCRS>.<Code>.Value = Nothing Then
                    NewCompoundCRS.HorizontalCRS.Code = 0
                Else
                    NewCompoundCRS.HorizontalCRS.Code = crsItem.<HorizontalCRS>.<Code>.Value
                End If
                NewCompoundCRS.HorizontalCRS.Comments = crsItem.<HorizontalCRS>.<Comments>.Value
                NewCompoundCRS.HorizontalCRS.Scope = crsItem.<HorizontalCRS>.<Scope>.Value
                NewCompoundCRS.HorizontalCRS.Deprecated = crsItem.<HorizontalCRS>.<Deprecated>.Value
                Dim HorizontalCRSAliasNames = From item In crsItem.<HorizontalCRS>.<AliasNames>.<AliasName>

                For Each nameItem In HorizontalCRSAliasNames
                    NewCompoundCRS.HorizontalCRS.AddAlias(nameItem)
                Next

                Select Case crsItem.<HorizontalCRS>.<Type>.Value
                    Case "Compound"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Compound
                    Case "Engineering"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Engineering
                    Case "Geocentric"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Geocentric
                    Case "Geographic2D"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Geographic2D
                    Case "Geographic3D"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Geographic3D
                    Case "Projected"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Projected
                    Case "Vertical"
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Vertical
                    Case Else
                        NewCompoundCRS.HorizontalCRS.Type = CrsTypes.Unknown
                End Select

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewCompoundCRS.Area.Name = crsItem.<Area>.<Name>.Value
                NewCompoundCRS.Area.Author = crsItem.<Area>.<Author>.Value
                NewCompoundCRS.Area.Code = crsItem.<Area>.<Code>.Value

                NewCompoundCRS.HorizontalCRS.CoordinateSystem.Name = crsItem.<HorizontalCRS>.<CoordinateSystem>.<Name>.Value
                NewCompoundCRS.HorizontalCRS.CoordinateSystem.Author = crsItem.<HorizontalCRS>.<CoordinateSystem>.<Author>.Value
                NewCompoundCRS.HorizontalCRS.CoordinateSystem.Code = crsItem.<HorizontalCRS>.<CoordinateSystem>.<Code>.Value
                NewCompoundCRS.HorizontalCRS.CoordinateSystem.Type = crsItem.<HorizontalCRS>.<CoordinateSystem>.<Type>.Value

                NewCompoundCRS.HorizontalCRS.Area.Name = crsItem.<HorizontalCRS>.<Area>.<Name>.Value
                NewCompoundCRS.HorizontalCRS.Area.Author = crsItem.<HorizontalCRS>.<Area>.<Author>.Value
                NewCompoundCRS.HorizontalCRS.Area.Code = crsItem.<HorizontalCRS>.<Area>.<Code>.Value


                'Read Vertical CRS information: ------------------------------------------------------------------------------------------------
                NewCompoundCRS.VerticalCRS.Name = crsItem.<VerticalCRS>.<Name>.Value
                NewCompoundCRS.VerticalCRS.Author = crsItem.<VerticalCRS>.<Author>.Value
                If crsItem.<VerticalCRS>.<Code>.Value = Nothing Then
                    NewCompoundCRS.VerticalCRS.Code = 0
                Else
                    NewCompoundCRS.VerticalCRS.Code = crsItem.<VerticalCRS>.<Code>.Value
                End If
                NewCompoundCRS.VerticalCRS.Comments = crsItem.<VerticalCRS>.<Comments>.Value
                NewCompoundCRS.VerticalCRS.Scope = crsItem.<VerticalCRS>.<Scope>.Value
                NewCompoundCRS.VerticalCRS.Deprecated = crsItem.<VerticalCRS>.<Deprecated>.Value
                Dim VerticalCRSAliasNames = From item In crsItem.<VerticalCRS>.<AliasNames>.<AliasName>

                For Each nameItem In VerticalCRSAliasNames
                    NewCompoundCRS.VerticalCRS.AddAlias(nameItem)
                Next

                NewCompoundCRS.VerticalCRS.CoordinateSystem.Name = crsItem.<VerticalCRS>.<CoordinateSystem>.<Name>.Value
                NewCompoundCRS.VerticalCRS.CoordinateSystem.Author = crsItem.<VerticalCRS>.<CoordinateSystem>.<Author>.Value
                NewCompoundCRS.VerticalCRS.CoordinateSystem.Code = crsItem.<VerticalCRS>.<CoordinateSystem>.<Code>.Value
                NewCompoundCRS.VerticalCRS.CoordinateSystem.Type = crsItem.<VerticalCRS>.<CoordinateSystem>.<Type>.Value

                NewCompoundCRS.VerticalCRS.Datum.Name = crsItem.<VerticalCRS>.<Datum>.<Name>.Value
                NewCompoundCRS.VerticalCRS.Datum.Author = crsItem.<VerticalCRS>.<Datum>.<Author>.Value
                NewCompoundCRS.VerticalCRS.Datum.Code = crsItem.<VerticalCRS>.<Datum>.<Code>.Value

                NewCompoundCRS.VerticalCRS.Area.Name = crsItem.<VerticalCRS>.<Area>.<Name>.Value
                NewCompoundCRS.VerticalCRS.Area.Author = crsItem.<VerticalCRS>.<Area>.<Author>.Value
                NewCompoundCRS.VerticalCRS.Area.Code = crsItem.<VerticalCRS>.<Area>.<Code>.Value


               

                List.Add(NewCompoundCRS)
            Next
        End Sub

        'Load the CRS list from the EPSG database.
        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'All CRSs will be read
            'da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System] WHERE COORD_REF_SYS_KIND = 'vertical'", myConnection)
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System]", myConnection)
            TableName = "CoordRefSys"
            da.Fill(ds, TableName)
            'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
            'CMPD_HORIZCRS_CODE and CMPD_VERTCRS_CODE are only used for compound CRS types

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)
            'AREA_CODE AREA_NAME AREA_OF_USE AREA_SOUTH_BOUND_LAT AREA_NORTH_BOUND_LAT AREA_WEST_BOUND_LON AREA_EAST_BOUND_LON AREA_POLYGON_FILE_REF ISO_A2_CODE ISO_A3_CODE ISO_N_CODE REMARKS DEPRECATED

            'Read the Coordinate System table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate System]", myConnection)
            TableName = "CoordSys"
            da.Fill(ds, TableName)
            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED

            'Read the Datum table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Datum]", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)
            'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            'Dim AreaOfUseCode As Integer
            'Dim CoordSysCode As Integer
            'Dim TargetUOMCode As Integer
            'Dim DatumCode As Integer
            'Dim ProjectionConvCode As Integer

            Dim NRows As Integer = ds.Tables("CoordRefSys").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Compound CRS list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Compound CRSs from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordRefSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                If ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_KIND") = "compound" Then 'Add this record to the list:

                    Dim NewCRS As New CompoundCRSSummary

                    NewCRS.Name = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_NAME")
                    NewCRS.Author = "EPSG"
                    NewCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_CODE")
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")) Then
                        NewCRS.Comments = ""
                    Else
                        NewCRS.Comments = ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")
                    End If

                    NewCRS.Selected = False
                    NewCRS.Scope = ds.Tables("CoordRefSys").Rows(RowNo).Item("CRS_SCOPE")
                    NewCRS.Deprecated = ds.Tables("CoordRefSys").Rows(RowNo).Item("DEPRECATED")

                    'Add list of alias names --------------------------------------------------------------------------------------------
                    'expression = "[OBJECT_TABLE_NAME] = 'Prime Meridian' AND [OBJECT_CODE] = " & Str(NewDatum.PrimeMeridian.Code)
                    expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.Code)
                    'Dim pmAliasResult = ds.Tables("Alias").Select(expression)
                    Dim result = ds.Tables("Alias").Select(expression)
                    For Each item In result
                        NewCRS.AddAlias(item.Item("ALIAS").ToString)
                    Next

                    'Add Area of Use information -----------------------------------------------------------------------------------------
                    NewCRS.Area.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("AREA_OF_USE_CODE")
                    expression = "[AREA_CODE] = " & Str(NewCRS.Area.Code)
                    Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                    If areaOfUseParameters.Count > 0 Then
                        NewCRS.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                        NewCRS.Area.Author = "EPSG"
                    End If

                    'Get Horizontal CRS information -------------------------------------------------------------------------------------
                    'CMPD_HORIZCRS_CODE and CMPD_VERTCRS_CODE are only used for compound CRS types
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("CMPD_HORIZCRS_CODE")) Then
                        RaiseEvent ErrorMessage("No horizontal CRS defined for Compound CRS named: " & NewCRS.Name & vbCrLf)
                    Else
                        NewCRS.HorizontalCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("CMPD_HORIZCRS_CODE")
                        expression = "[COORD_REF_SYS_CODE] = " & Str(NewCRS.HorizontalCRS.Code)
                        Dim HorizontalCrs = ds.Tables("CoordRefSys").Select(expression)
                        'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
                        If HorizontalCrs.Count > 0 Then
                            NewCRS.HorizontalCRS.Name = HorizontalCrs(0).Item("COORD_REF_SYS_NAME")
                            NewCRS.HorizontalCRS.Author = "EPSG"
                            If IsDBNull(HorizontalCrs(0).Item("REMARKS")) Then
                                NewCRS.HorizontalCRS.Comments = ""
                            Else
                                NewCRS.HorizontalCRS.Comments = HorizontalCrs(0).Item("REMARKS")
                            End If
                            NewCRS.HorizontalCRS.Scope = HorizontalCrs(0).Item("CRS_SCOPE")
                            NewCRS.HorizontalCRS.Deprecated = HorizontalCrs(0).Item("DEPRECATED")
                            expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.HorizontalCRS.Code)
                            Dim result2 = ds.Tables("Alias").Select(expression)
                            For Each item In result2
                                NewCRS.HorizontalCRS.AddAlias(item.Item("ALIAS").ToString)
                            Next

                            Select Case HorizontalCrs(0).Item("COORD_REF_SYS_KIND")
                                Case "compound"
                                    NewCRS.HorizontalCRS.Type = CrsTypes.Compound
                                Case "engineering"
                                    NewCRS.HorizontalCRS.Type = CrsTypes.Engineering
                                Case "geocentric"
                                    NewCRS.HorizontalCRS.Type = CrsTypes.Geocentric
                                Case "geographic 2D"
                                    NewCRS.HorizontalCRS.Type = CrsTypes.Geographic2D
                                Case "geographic 3D"
                                    NewCRS.HorizontalCRS.Type = CrsTypes.Geographic3D
                                Case "projected"
                                    NewCRS.HorizontalCRS.Type = CrsTypes.Projected
                                Case "vertical"
                                    NewCRS.HorizontalCRS.Type = CrsTypes.Vertical
                                Case Else
                                    NewCRS.HorizontalCRS.Type = CrsTypes.Unknown
                            End Select

                            NewCRS.HorizontalCRS.CoordinateSystem.Code = HorizontalCrs(0).Item("COORD_SYS_CODE")
                            NewCRS.HorizontalCRS.CoordinateSystem.Author = "EPSG"
                            expression = "[COORD_SYS_CODE] = " & Str(NewCRS.HorizontalCRS.CoordinateSystem.Code)
                            Dim coordSysParameters2 = ds.Tables("CoordSys").Select(expression)
                            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                            If coordSysParameters2.Count > 0 Then
                                NewCRS.HorizontalCRS.CoordinateSystem.Name = coordSysParameters2(0).Item("COORD_SYS_NAME").ToString
                                NewCRS.HorizontalCRS.CoordinateSystem.Type = coordSysParameters2(0).Item("COORD_SYS_TYPE").ToString
                            End If

                            NewCRS.HorizontalCRS.Area.Code = HorizontalCrs(0).Item("AREA_OF_USE_CODE")
                            NewCRS.HorizontalCRS.Area.Author = "EPSG"
                            expression = "[AREA_CODE] = " & Str(NewCRS.HorizontalCRS.Area.Code)
                            Dim areaOfUseParameters2 = ds.Tables("Area").Select(expression)
                            If areaOfUseParameters2.Count > 0 Then
                                NewCRS.HorizontalCRS.Area.Name = areaOfUseParameters2(0).Item("AREA_NAME").ToString
                            End If
                        End If
                    End If
                 

                    'Get Horizontal CRS information -------------------------------------------------------------------------------------
                    'CMPD_HORIZCRS_CODE and CMPD_VERTCRS_CODE are only used for compound CRS types
                    NewCRS.VerticalCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("CMPD_VERTCRS_CODE")
                    expression = "[COORD_REF_SYS_CODE] = " & Str(NewCRS.VerticalCRS.Code)
                    Dim VerticalCRS = ds.Tables("CoordRefSys").Select(expression)
                    'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
                    If VerticalCRS.Count > 0 Then
                        NewCRS.VerticalCRS.Name = VerticalCRS(0).Item("COORD_REF_SYS_NAME")
                        NewCRS.VerticalCRS.Author = "EPSG"
                        If IsDBNull(VerticalCRS(0).Item("REMARKS")) Then
                            NewCRS.VerticalCRS.Comments = ""
                        Else
                            NewCRS.VerticalCRS.Comments = VerticalCRS(0).Item("REMARKS")
                        End If
                        NewCRS.VerticalCRS.Scope = VerticalCRS(0).Item("CRS_SCOPE")
                        NewCRS.VerticalCRS.Deprecated = VerticalCRS(0).Item("DEPRECATED")

                        expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.VerticalCRS.Code)
                        Dim result2 = ds.Tables("Alias").Select(expression)
                        For Each item In result2
                            NewCRS.VerticalCRS.AddAlias(item.Item("ALIAS").ToString)
                        Next

                        NewCRS.VerticalCRS.CoordinateSystem.Code = VerticalCRS(0).Item("COORD_SYS_CODE")
                        NewCRS.VerticalCRS.CoordinateSystem.Author = "EPSG"
                        expression = "[COORD_SYS_CODE] = " & Str(NewCRS.VerticalCRS.CoordinateSystem.Code)
                        Dim coordSysParameters3 = ds.Tables("CoordSys").Select(expression)
                        'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                        If coordSysParameters3.Count > 0 Then
                            NewCRS.VerticalCRS.CoordinateSystem.Name = coordSysParameters3(0).Item("COORD_SYS_NAME").ToString
                            NewCRS.VerticalCRS.CoordinateSystem.Type = coordSysParameters3(0).Item("COORD_SYS_TYPE").ToString
                        End If

                        NewCRS.VerticalCRS.Datum.Code = VerticalCRS(0).Item("DATUM_CODE")
                        NewCRS.VerticalCRS.Datum.Author = "EPSG"
                        expression = "[DATUM_CODE] = " & Str(NewCRS.VerticalCRS.Datum.Code)
                        'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED
                        Dim datumParameters = ds.Tables("Datum").Select(expression)
                        If datumParameters.Count > 0 Then
                            NewCRS.VerticalCRS.Datum.Name = datumParameters(0).Item("DATUM_NAME")
                            NewCRS.VerticalCRS.Datum.Type = datumParameters(0).Item("DATUM_TYPE")
                        End If

                        NewCRS.VerticalCRS.Area.Code = VerticalCRS(0).Item("AREA_OF_USE_CODE")
                        NewCRS.VerticalCRS.Area.Author = "EPSG"
                        expression = "[AREA_CODE] = " & Str(NewCRS.VerticalCRS.Area.Code)
                        Dim areaOfUseParameters3 = ds.Tables("Area").Select(expression)
                        If areaOfUseParameters3.Count > 0 Then
                            NewCRS.VerticalCRS.Area.Name = areaOfUseParameters3(0).Item("AREA_NAME").ToString
                        End If

                    End If




                    ''Add Coordinate System information -----------------------------------------------------------------------------------------
                    ''Dim coordSystem As New XElement("CoordinateSystem")
                    'NewCRS.CoordinateSystem.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_SYS_CODE")
                    'expression = "[COORD_SYS_CODE] = " & Str(NewCRS.CoordinateSystem.Code)
                    'Dim coordSysParameters = ds.Tables("CoordSys").Select(expression)
                    ''COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                    'If coordSysParameters.Count > 0 Then
                    '    NewCRS.CoordinateSystem.Name = coordSysParameters(0).Item("COORD_SYS_NAME").ToString
                    '    NewCRS.CoordinateSystem.Author = "EPSG"
                    '    NewCRS.CoordinateSystem.Type = coordSysParameters(0).Item("COORD_SYS_TYPE").ToString
                    'End If

                    ''Add Datum information -----------------------------------------------------------------------------------------
                    'If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")) Then
                    '    NewCRS.Datum.Code = 0
                    'Else
                    '    NewCRS.Datum.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")
                    'End If

                    'expression = "[DATUM_CODE] = " & Str(NewCRS.Datum.Code)
                    'Dim datumParameters = ds.Tables("Datum").Select(expression)
                    ''DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED
                    'If datumParameters.Count > 0 Then
                    '    NewCRS.Datum.Name = datumParameters(0).Item("DATUM_NAME")
                    '    NewCRS.Datum.Author = "EPSG"
                    '    NewCRS.Datum.Type = datumParameters(0).Item("DATUM_TYPE")
                    'End If

                 
                    List.Add(NewCRS)
                End If
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Compound CRSs as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Compound Coordinate Reference System List File-->
                       <CompoundCRSList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <CompoundCRS>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <Area>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </Area>
                                     <HorizontalCRS>
                                         <Name><%= item.HorizontalCRS.Name %></Name>
                                         <Author><%= item.HorizontalCRS.Author %></Author>
                                         <Code><%= item.HorizontalCRS.Code %></Code>
                                         <Comments><%= item.HorizontalCRS.Comments %></Comments>
                                         <Scope><%= item.HorizontalCRS.Scope %></Scope>
                                         <Deprecated><%= item.HorizontalCRS.Deprecated %></Deprecated>
                                         <AliasNames>
                                             <%= From nameItem In item.HorizontalCRS.AliasName _
                                                 Select _
                                                 <AliasName><%= nameItem %></AliasName>
                                             %>
                                         </AliasNames>
                                         <Type><%= item.HorizontalCRS.Type %></Type>
                                         <CoordinateSystem>
                                             <Name><%= item.HorizontalCRS.CoordinateSystem.Name %></Name>
                                             <Author><%= item.HorizontalCRS.CoordinateSystem.Author %></Author>
                                             <Code><%= item.HorizontalCRS.CoordinateSystem.Code %></Code>
                                             <Type><%= item.HorizontalCRS.CoordinateSystem.Type %></Type>
                                         </CoordinateSystem>
                                         <Area>
                                             <Name><%= item.HorizontalCRS.Area.Name %></Name>
                                             <Author><%= item.HorizontalCRS.Area.Author %></Author>
                                             <Code><%= item.HorizontalCRS.Area.Code %></Code>
                                         </Area>
                                     </HorizontalCRS>
                                     <VerticalCRS>
                                         <Name><%= item.VerticalCRS.Name %></Name>
                                         <Author><%= item.VerticalCRS.Author %></Author>
                                         <Code><%= item.VerticalCRS.Code %></Code>
                                         <Comments><%= item.VerticalCRS.Comments %></Comments>
                                         <Scope><%= item.VerticalCRS.Scope %></Scope>
                                         <Deprecated><%= item.VerticalCRS.Deprecated %></Deprecated>
                                         <AliasNames>
                                             <%= From nameItem In item.VerticalCRS.AliasName _
                                                 Select _
                                                 <AliasName><%= nameItem %></AliasName>
                                             %>
                                         </AliasNames>
                                         <CoordinateSystem>
                                             <Name><%= item.VerticalCRS.CoordinateSystem.Name %></Name>
                                             <Author><%= item.VerticalCRS.CoordinateSystem.Author %></Author>
                                             <Code><%= item.VerticalCRS.CoordinateSystem.Code %></Code>
                                             <Type><%= item.VerticalCRS.CoordinateSystem.Type %></Type>
                                         </CoordinateSystem>
                                         <Datum>
                                             <Name><%= item.VerticalCRS.Datum.Name %></Name>
                                             <Author><%= item.VerticalCRS.Datum.Author %></Author>
                                             <Code><%= item.VerticalCRS.Datum.Code %></Code>
                                             <Type><%= item.VerticalCRS.Datum.Type %></Type>
                                         </Datum>
                                         <Area>
                                             <Name><%= item.VerticalCRS.Area.Name %></Name>
                                             <Author><%= item.VerticalCRS.Area.Author %></Author>
                                             <Code><%= item.VerticalCRS.Area.Code %></Code>
                                         </Area>
                                     </VerticalCRS>
                                 </CompoundCRS>
                           %>
                           <!---->
                       </CompoundCRSList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------


#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'CompoundCRSList ---------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class EngineeringCRSList '---------------------------------------------------------------------------------------------------------------------------------------------------------
        Public List As New List(Of EngineeringCRSSummary)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property
#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Engineering CRS List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<EngineeringCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<EngineeringCRSList>.<LastEditDate>.Value
            Description = XDoc.<EngineeringCRSList>.<Description>.Value

            Dim EngCRSs = From item In XDoc.<EngineeringCRSList>.<EngineeringCRS>

            List.Clear()
            For Each crsItem In EngCRSs
                Dim NewEngCRS As New EngineeringCRSSummary
                NewEngCRS.Name = crsItem.<Name>.Value
                NewEngCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewEngCRS.Code = 0
                Else
                    NewEngCRS.Code = crsItem.<Code>.Value
                End If
                NewEngCRS.Selected = crsItem.<Selected>.Value
                NewEngCRS.Comments = crsItem.<Comments>.Value
                NewEngCRS.Scope = crsItem.<Scope>.Value
                NewEngCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewEngCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewEngCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewEngCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewEngCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewEngCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewEngCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewEngCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewEngCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewEngCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewEngCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewEngCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewEngCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewEngCRS)
            Next
        End Sub

        'Load the list from the selected list file.
        Public Sub LoadFile()
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<EngineeringCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<EngineeringCRSList>.<LastEditDate>.Value
            Description = XDoc.<EngineeringCRSList>.<Description>.Value

            Dim EngCRSs = From item In XDoc.<EngineeringCRSList>.<EngineeringCRS>

            List.Clear()
            For Each crsItem In EngCRSs
                Dim NewEngCRS As New EngineeringCRSSummary
                NewEngCRS.Name = crsItem.<Name>.Value
                NewEngCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewEngCRS.Code = 0
                Else
                    NewEngCRS.Code = crsItem.<Code>.Value
                End If
                NewEngCRS.Selected = crsItem.<Selected>.Value
                NewEngCRS.Comments = crsItem.<Comments>.Value
                NewEngCRS.Scope = crsItem.<Scope>.Value
                NewEngCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewEngCRS.AddAlias(nameItem)
                Next

                'Read Datum information: -------------------------------------------------------------------------------------------------
                NewEngCRS.Datum.Name = crsItem.<Datum>.<Name>.Value
                NewEngCRS.Datum.Author = crsItem.<Datum>.<Author>.Value
                NewEngCRS.Datum.Code = crsItem.<Datum>.<Code>.Value
                NewEngCRS.Datum.Type = crsItem.<Datum>.<Type>.Value

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewEngCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewEngCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewEngCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewEngCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewEngCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewEngCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewEngCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewEngCRS)
            Next
        End Sub

        'Load the CRS list from the EPSG database.
        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Only records where COORD_REF_SYS_KIND = 'engineering' will be processed.)
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System] WHERE COORD_REF_SYS_KIND = 'engineering'", myConnection)
            TableName = "CoordRefSys"
            da.Fill(ds, TableName)
            'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
            'CMPD_HORIZCRS_CODE and CMPD_VERTCRS_CODE are only used for compound CRS types

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)
            'AREA_CODE AREA_NAME AREA_OF_USE AREA_SOUTH_BOUND_LAT AREA_NORTH_BOUND_LAT AREA_WEST_BOUND_LON AREA_EAST_BOUND_LON AREA_POLYGON_FILE_REF ISO_A2_CODE ISO_A3_CODE ISO_N_CODE REMARKS DEPRECATED

            'Read the Coordinate System table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate System]", myConnection)
            TableName = "CoordSys"
            da.Fill(ds, TableName)
            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED

            'Read the Datum table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Datum]", myConnection)
            TableName = "Datum"
            da.Fill(ds, TableName)
            'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED

            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number
            'Dim AreaOfUseCode As Integer 'The Area of Use code number
            'Dim AreaOfUseCode As Integer
            'Dim CoordSysCode As Integer
            Dim TargetUOMCode As Integer
            'Dim DatumCode As Integer
            'Dim ProjectionConvCode As Integer

            Dim NRows As Integer = ds.Tables("CoordRefSys").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Engineering CRS list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Engineering CRSs from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordRefSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                'If ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_KIND") = "geographic 2D" Then 'Add this record to the list:
                Dim NewCRS As New EngineeringCRSSummary
                NewCRS.Name = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_NAME")
                NewCRS.Author = "EPSG"
                NewCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_CODE")
                If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")) Then
                    NewCRS.Comments = ""
                Else
                    NewCRS.Comments = ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")
                End If

                NewCRS.Scope = ds.Tables("CoordRefSys").Rows(RowNo).Item("CRS_SCOPE")
                NewCRS.Deprecated = ds.Tables("CoordRefSys").Rows(RowNo).Item("DEPRECATED")

                'Add list of alias names --------------------------------------------------------------------------------------------
                'expression = "[OBJECT_TABLE_NAME] = 'Prime Meridian' AND [OBJECT_CODE] = " & Str(NewDatum.PrimeMeridian.Code)
                expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.Code)
                'Dim pmAliasResult = ds.Tables("Alias").Select(expression)
                Dim result = ds.Tables("Alias").Select(expression)
                For Each item In result
                    NewCRS.AddAlias(item.Item("ALIAS").ToString)
                Next

                'Add Coordinate System information -----------------------------------------------------------------------------------------
                'Dim coordSystem As New XElement("CoordinateSystem")
                NewCRS.CoordinateSystem.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_SYS_CODE")
                expression = "[COORD_SYS_CODE] = " & Str(NewCRS.CoordinateSystem.Code)
                Dim coordSysParameters = ds.Tables("CoordSys").Select(expression)
                'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                If coordSysParameters.Count > 0 Then
                    NewCRS.CoordinateSystem.Name = coordSysParameters(0).Item("COORD_SYS_NAME").ToString
                    NewCRS.CoordinateSystem.Author = "EPSG"
                    NewCRS.CoordinateSystem.Type = coordSysParameters(0).Item("COORD_SYS_TYPE").ToString
                End If

                'Add Datum information -----------------------------------------------------------------------------------------
                If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")) Then
                    NewCRS.Datum.Code = 0
                Else
                    NewCRS.Datum.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("DATUM_CODE")
                End If

                expression = "[DATUM_CODE] = " & Str(NewCRS.Datum.Code)
                Dim datumParameters = ds.Tables("Datum").Select(expression)
                'DATUM_CODE DATUM_NAME DATUM_TYPE ORIGIN_DESCRIPTION REALIZATION_EPOCH ELLIPSOID_CODE PRIME_MERIDIAN_CODE AREA_OF_USE_CODE DATUM_SCOPE REMARKS DEPRECATED
                If datumParameters.Count > 0 Then
                    NewCRS.Datum.Name = datumParameters(0).Item("DATUM_NAME")
                    NewCRS.Datum.Author = "EPSG"
                    NewCRS.Datum.Type = datumParameters(0).Item("DATUM_TYPE")
                End If

                'Add Area of Use information -----------------------------------------------------------------------------------------
                NewCRS.Area.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("AREA_OF_USE_CODE")
                expression = "[AREA_CODE] = " & Str(NewCRS.Area.Code)
                Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                If areaOfUseParameters.Count > 0 Then
                    NewCRS.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                    NewCRS.Area.Author = "EPSG"
                End If
                List.Add(NewCRS)
                'End If
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Engineering CRS as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Engineering Coordinate Reference System List File-->
                       <EngineeringCRSList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <EngineeringCRS>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <Datum>
                                         <Name><%= item.Datum.Name %></Name>
                                         <Author><%= item.Datum.Author %></Author>
                                         <Code><%= item.Datum.Code %></Code>
                                         <Type><%= item.Datum.Type %></Type>
                                     </Datum>
                                     <CoordinateSystem>
                                         <Name><%= item.CoordinateSystem.Name %></Name>
                                         <Author><%= item.CoordinateSystem.Author %></Author>
                                         <Code><%= item.CoordinateSystem.Code %></Code>
                                         <Type><%= item.CoordinateSystem.Type %></Type>
                                     </CoordinateSystem>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </EngineeringCRS>
                           %>
                           <!---->
                       </EngineeringCRSList>
            Return XDoc
        End Function

        Public Sub AddUser()
            'Add a user to the list
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        Public Sub RemoveUser()
            'Remove a user for the list.
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'EngineeringCRSList ------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Class ProjectedCRSList '-----------------------------------------------------------------------------------------------------------------------------------------------------------

        Public List As New List(Of ProjectedCRS)
    Public FileLocation As New ADVL_Utilities_Library_1.FileLocation

#Region " Properties" '---------------------------------------------------------------------------------------------------

    Private _listFileName As String = ""
        Property ListFileName As String 'The file name (with extension) of the list file.
            Get
                Return _listFileName
            End Get
            Set(value As String)
                _listFileName = value
            End Set
        End Property

        Private _creationDate As DateTime = Now
        Property CreationDate As DateTime
            Get
                Return _creationDate
            End Get
            Set(value As DateTime)
                _creationDate = value
            End Set
        End Property

        Private _lastEditDate As DateTime = Now
        Property LastEditDate As DateTime
            Get
                Return _lastEditDate
            End Get
            Set(value As DateTime)
                _lastEditDate = value
            End Set
        End Property


        Private _description As String = ""
        Property Description As String
            Get
                Return _description
            End Get
            Set(value As String)
                _description = value
            End Set
        End Property

        Private _nRecords As Integer = 0 'The number of record in the list
        ReadOnly Property NRecords As Integer
            Get
                _nRecords = List.Count
                Return _nRecords
            End Get
        End Property

        Private _nUsers As Integer = 0
        Property NUsers As Integer 'The number of users connected ot the list.
            Get
                Return _nUsers
            End Get
            Set(value As Integer)
                _nUsers = value
            End Set
        End Property

#End Region 'Properties --------------------------------------------------------------------------------------------------

#Region "Methods" '-------------------------------------------------------------------------------------------------------

        'Clear the list.
        Public Sub Clear()
            List.Clear()
        FileLocation.Type = ADVL_Utilities_Library_1.FileLocation.Types.Directory
        FileLocation.Path = ""
            ListFileName = ""
            Description = ""
            NUsers = 0
        End Sub

        'Load the XML data in the XDoc into the Projected CRS List.
        Public Sub LoadXml(ByRef XDoc As System.Xml.Linq.XDocument)

            CreationDate = XDoc.<ProjectedCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<ProjectedCRSList>.<LastEditDate>.Value
            Description = XDoc.<ProjectedCRSList>.<Description>.Value

            Dim ProjectedCrsList = From item In XDoc.<ProjectedCRSList>.<ProjectedCRS>

            List.Clear()
            For Each crsItem In ProjectedCrsList
                Dim NewCRS As New ProjectedCRS
                NewCRS.Name = crsItem.<Name>.Value
                NewCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewCRS.Code = 0
                Else
                    NewCRS.Code = crsItem.<Code>.Value
                End If
                NewCRS.Selected = crsItem.<Selected>.Value
                NewCRS.Comments = crsItem.<Comments>.Value
                NewCRS.Scope = crsItem.<Scope>.Value
                NewCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewCRS.AddAlias(nameItem)
                Next

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read SourceGeographicCRS information: -------------------------------------------------------------------------------------------------
                NewCRS.SourceGeographicCRS.Name = crsItem.<SourceGeographicCRS>.<Name>.Value
                NewCRS.SourceGeographicCRS.Author = crsItem.<SourceGeographicCRS>.<Author>.Value
                NewCRS.SourceGeographicCRS.Code = crsItem.<SourceGeographicCRS>.<Code>.Value
                NewCRS.SourceGeographicCRS.Type = crsItem.<SourceGeographicCRS>.<Type>.Value

                'Read Projection information ---------------------------------------------------------------------------------------------------
                NewCRS.Projection.Name = crsItem.<Projection>.<Name>.Value
                NewCRS.Projection.Author = crsItem.<Projection>.<Author>.Value
                NewCRS.Projection.Code = crsItem.<Projection>.<Code>.Value
                NewCRS.Projection.Type = crsItem.<Projection>.<Type>.Value

                'Read Projection Method information --------------------------------------------------------------------------------------------
                NewCRS.ProjectionMethod.Name = crsItem.<ProjectionMethod>.<Name>.Value
                NewCRS.ProjectionMethod.Author = crsItem.<ProjectionMethod>.<Author>.Value
                NewCRS.ProjectionMethod.Code = crsItem.<ProjectionMethod>.<Code>.Value
                NewCRS.ProjectionMethod.Type = crsItem.<ProjectionMethod>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewCRS)
            Next
        End Sub

        'Load the list from the selected list file.
        Public Sub LoadFile()
            If ListFileName = "" Then 'No list file has been selected.
                Exit Sub
            End If

            Dim XDoc As System.Xml.Linq.XDocument
            FileLocation.ReadXmlData(ListFileName, XDoc)

            If XDoc Is Nothing Then
                RaiseEvent ErrorMessage("Xml list file is blank." & vbCrLf)
                Exit Sub
            End If

            CreationDate = XDoc.<ProjectedCRSList>.<CreationDate>.Value
            LastEditDate = XDoc.<ProjectedCRSList>.<LastEditDate>.Value
            Description = XDoc.<ProjectedCRSList>.<Description>.Value

            Dim ProjectedCrsList = From item In XDoc.<ProjectedCRSList>.<ProjectedCRS>

            List.Clear()
            For Each crsItem In ProjectedCrsList
                Dim NewCRS As New ProjectedCRS
                NewCRS.Name = crsItem.<Name>.Value
                NewCRS.Author = crsItem.<Author>.Value
                If crsItem.<Code>.Value = Nothing Then
                    NewCRS.Code = 0
                Else
                    NewCRS.Code = crsItem.<Code>.Value
                End If
                NewCRS.Selected = crsItem.<Selected>.Value
                NewCRS.Comments = crsItem.<Comments>.Value
                NewCRS.Scope = crsItem.<Scope>.Value
                NewCRS.Deprecated = crsItem.<Deprecated>.Value

                Dim aliasNames = From item In crsItem.<AliasNames>.<AliasName>
                For Each nameItem In aliasNames
                    NewCRS.AddAlias(nameItem)
                Next

                'Read CoordinateSystem information: -------------------------------------------------------------------------------------------------
                NewCRS.CoordinateSystem.Name = crsItem.<CoordinateSystem>.<Name>.Value
                NewCRS.CoordinateSystem.Author = crsItem.<CoordinateSystem>.<Author>.Value
                NewCRS.CoordinateSystem.Code = crsItem.<CoordinateSystem>.<Code>.Value
                NewCRS.CoordinateSystem.Type = crsItem.<CoordinateSystem>.<Type>.Value

                'Read SourceGeographicCRS information: -------------------------------------------------------------------------------------------------
                NewCRS.SourceGeographicCRS.Name = crsItem.<SourceGeographicCRS>.<Name>.Value
                NewCRS.SourceGeographicCRS.Author = crsItem.<SourceGeographicCRS>.<Author>.Value
                NewCRS.SourceGeographicCRS.Code = crsItem.<SourceGeographicCRS>.<Code>.Value
                NewCRS.SourceGeographicCRS.Type = crsItem.<SourceGeographicCRS>.<Type>.Value

                'Read Projection information ---------------------------------------------------------------------------------------------------
                NewCRS.Projection.Name = crsItem.<Projection>.<Name>.Value
                NewCRS.Projection.Author = crsItem.<Projection>.<Author>.Value
                NewCRS.Projection.Code = crsItem.<Projection>.<Code>.Value
                NewCRS.Projection.Type = crsItem.<Projection>.<Type>.Value

                'Read Projection Method information --------------------------------------------------------------------------------------------
                NewCRS.ProjectionMethod.Name = crsItem.<ProjectionMethod>.<Name>.Value
                NewCRS.ProjectionMethod.Author = crsItem.<ProjectionMethod>.<Author>.Value
                NewCRS.ProjectionMethod.Code = crsItem.<ProjectionMethod>.<Code>.Value
                NewCRS.ProjectionMethod.Type = crsItem.<ProjectionMethod>.<Type>.Value

                'Read Area of Use information: -------------------------------------------------------------------------------------------------
                NewCRS.Area.Name = crsItem.<AreaOfUse>.<Name>.Value
                NewCRS.Area.Author = crsItem.<AreaOfUse>.<Author>.Value
                NewCRS.Area.Code = crsItem.<AreaOfUse>.<Code>.Value

                List.Add(NewCRS)
            Next

        End Sub

        'Load the CRS list from the EPSG database.
        Public Sub LoadEpsgDbList(ByVal EpsgDatabasePath As String)
            If EpsgDatabasePath = "" Then 'No EPSG database path has been selected.
                RaiseEvent ErrorMessage("EPSG database path has not been selected." & vbCrLf)
                Exit Sub
            End If

            If System.IO.File.Exists(EpsgDatabasePath) = False Then 'Database not found.
                RaiseEvent ErrorMessage("EPSG database path is not valid." & vbCrLf)
                Exit Sub
            End If

            Dim connString As String
            Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
            Dim ds As DataSet = New DataSet
            Dim da As OleDb.OleDbDataAdapter
            Dim tables As DataTableCollection = ds.Tables

            Dim TableName As String = ""

            connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & EpsgDatabasePath
            myConnection.ConnectionString = connString


            'Check if a database lock file exists: ---------------------------------------------------------------------------------------
            'Main.MessageAdd("looking for a lock file corresponding to the database path: " & Main.EpsgDatabasePath & vbCrLf)

            Dim LockFileName As String = System.IO.Path.GetFileNameWithoutExtension(EpsgDatabasePath) & ".ldb"
            Dim LockFilePath As String = System.IO.Path.GetDirectoryName(EpsgDatabasePath) & "\" & LockFileName
            'Main.MessageAdd("Lock file path: " & LockFilePath & vbCrLf)

            If System.IO.File.Exists(LockFilePath) Then
                'Main.MessageAdd("Database lock file found: " & LockFilePath & vbCrLf)
                'Main.MessageAdd("Database will not be opened!" & vbCrLf)
                RaiseEvent ErrorMessage("Database lock file found: " & LockFilePath & vbCrLf)
                RaiseEvent ErrorMessage("Database will not be opened!" & vbCrLf)
                Exit Sub
            End If
            '-----------------------------------------------------------------------------------------------------------------------------

            myConnection.Open()

            ds.Clear()
            ds.Reset()

            'Read the Coordinate Reference System table into dataset ds. 'Read all the records - this table is also used to find the Base Coordinate Reference System.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate Reference System]", myConnection)
            TableName = "CoordRefSys"
            da.Fill(ds, TableName)
            'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
            'These fields in the Coordinate Reference System table are not used for projected CRSs: DATUM_CODE CMPD_HORIZCRS_CODE CMPD_VERTCRS_CODE
            'For Projected CRS, COORD_REF_SYS_KIND = projected.

            'Read the Alias table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Alias]", myConnection)
            TableName = "Alias"
            da.Fill(ds, TableName)
            'ALIAS CODE OBJECT_TABLE_NAME OBJECT_CODE NAMING_SYSTEM_CODE ALIAS REMARKS

            'Read the Area Of Use table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Area]", myConnection)
            TableName = "Area"
            da.Fill(ds, TableName)
            'AREA_CODE AREA_NAME AREA_OF_USE AREA_SOUTH_BOUND_LAT AREA_NORTH_BOUND_LAT AREA_WEST_BOUND_LON AREA_EAST_BOUND_LON AREA_POLYGON_FILE_REF ISO_A2_CODE ISO_A3_CODE ISO_N_CODE REMARKS DEPRECATED

            'Read the Coordinate System table into dataset ds.
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate System]", myConnection)
            TableName = "CoordSys"
            da.Fill(ds, TableName)
            'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED

            'Read the Coordinate_Operation table into dataset ds 
            da = New OleDb.OleDbDataAdapter("Select * From [Coordinate_Operation]", myConnection)
            TableName = "CoordOp"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_CODE, COORD_OP_NAME, COORD_OP_TYPE, SOURCE_CRS_CODE, TARGET_CRS_CODE, COORD_TFM_VERSION, COORD_OP_VARIANT, AREA_OF_USE_CODE, COORD_OP_SCOPE,
            'COORD_OP_ACCURACY, COORD_OP_METHOD_CODE, UOM_CODE_SOURCE_COORD_DIFF, UOM_CODE_TARGET_COORD_DIFF, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE,
            'REVISION_DATE, SHOW_OPERATION, DEPRECATED
            'Note: COORD_OP_TYPE = "conversion" is a map projection
            'Map projections do not use SOURCE_CRS_CODE or TARGET_CRS_CODE or COORD_TFM_VERSION or COORD_OP_VARIANT 
            'COORD_OP_ACCURACY is defined as 0 for map projections. (They are considered exact by definition.)

            'Read the Coordinate Operation Method table into dataset ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Method]", myConnection)
            TableName = "CoordOpMethod"
            da.Fill(ds, TableName)
            'Table fields: COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, FORMULA, EXAMPLE, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            'Read the Coordinate Operation Parameter Usage table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter Usage]", myConnection)
            TableName = "CoordOpUsage"
            da.Fill(ds, TableName)
            'Coordinate_Operation Parameter Usage Table fields :
            'Table fields: COORD_OP_METHOD_CODE, PARAMETER_CODE, SORT_ORDER, PARAM_SIGN_REVERSAL

            'Read the Coordinate Operation Parameter table into dataset ds:
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter]", myConnection)
            TableName = "CoordOpParams"
            da.Fill(ds, TableName)
            'Coordinate_Operation Parameter Table fields: 
            'PARAMETER_CODE, PARAMETER_NAME, DESCRIPTION, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED

            'Read the Coordinate Operation Parameter Value table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Coordinate_Operation Parameter Value]", myConnection)
            TableName = "CoordOpValues"
            da.Fill(ds, TableName)
            'Coordinate_Operation Parameter Value Table fields:
            'COORD_OP_CODE, COORD_OP_METHOD_CODE, PARAMETER_CODE, PARAMETER_VALUE, PARAM_VALUE_FILE_REF, UOM_CODE

            'Read the Unit of Measure table into ds
            da.SelectCommand = New OleDb.OleDbCommand("Select * From [Unit of Measure]", myConnection)
            TableName = "UnitOfMeasure"
            da.Fill(ds, TableName)
            'Unit of Measure Table fields :
            'UOM_CODE, UNIT_OF_MEAS_NAME, UNIT_OF_MEAS_TYPE, TARGET_UOM_CODE, FACTOR_B, FACTOR_C, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED


            Dim expression As String

            ' variables:
            'Dim NParams As Integer 'The number of parameters used to define the projection.
            'Dim ParamNo As Integer 'The current parameter number.
            'Dim ParamCode As Integer 'The parameter code of the current parameter.
            'Dim DatumCode As Integer 'The datum code number


            Dim NRows As Integer = ds.Tables("CoordRefSys").Rows.Count

            List.Clear()

            CreationDate = Now
            LastEditDate = Now
            Description = "Projected CRS list. Source: EPSG database. Date loaded: " & Format(Now, "d-MMM-yyyy H:mm:ss") & " Database path: " & EpsgDatabasePath

            Dim RowNo As Integer
            RaiseEvent Message("Reading Projected CRS list from EPSG database" & vbCrLf)
            For RowNo = 0 To NRows - 1 'loop through each row in the CoordRefSys table:

                'Send a progress message when every 100th item is read:
                If RowNo Mod 100 = 0 Then
                    RaiseEvent Message("Reading item " & RowNo & " of " & NRows & vbCrLf)
                End If

                If ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_KIND") = "projected" Then 'Add this record to the list:
                    Dim NewCRS As New ProjectedCRS
                    NewCRS.Name = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_NAME")
                    NewCRS.Author = "EPSG"
                    NewCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_REF_SYS_CODE")
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")) Then
                        NewCRS.Comments = ""
                    Else
                        NewCRS.Comments = ds.Tables("CoordRefSys").Rows(RowNo).Item("REMARKS")
                    End If
                    NewCRS.Scope = ds.Tables("CoordRefSys").Rows(RowNo).Item("CRS_SCOPE")
                    NewCRS.Deprecated = ds.Tables("CoordRefSys").Rows(RowNo).Item("DEPRECATED")

                    'Add list of alias names --------------------------------------------------------------------------------------------
                    expression = "[OBJECT_TABLE_NAME] = 'Coordinate Reference System' AND [OBJECT_CODE] = " & Str(NewCRS.Code)
                    Dim result = ds.Tables("Alias").Select(expression)
                    For Each item In result
                        NewCRS.AddAlias(item.Item("ALIAS").ToString)
                    Next

                    'Add Coordinate System information -----------------------------------------------------------------------------------------
                    NewCRS.CoordinateSystem.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("COORD_SYS_CODE")
                    expression = "[COORD_SYS_CODE] = " & Str(NewCRS.CoordinateSystem.Code)
                    Dim coordSysParameters = ds.Tables("CoordSys").Select(expression)
                    'COORD_SYS_CODE COORD_SYS_NAME COORD_SYS_TYPE DIMENSION REMARKS DEPRECATED
                    If coordSysParameters.Count > 0 Then
                        NewCRS.CoordinateSystem.Name = coordSysParameters(0).Item("COORD_SYS_NAME").ToString
                        NewCRS.CoordinateSystem.Author = "EPSG"
                        NewCRS.CoordinateSystem.Type = coordSysParameters(0).Item("COORD_SYS_TYPE").ToString
                    End If

                    'Add Source Geographic Coordinate System information -----------------------------------------------------------------------------------------
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("SOURCE_GEOGCRS_CODE")) Then
                        NewCRS.SourceGeographicCRS.Code = 0
                    Else
                        NewCRS.SourceGeographicCRS.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("SOURCE_GEOGCRS_CODE")
                    End If

                    'expression = "[SOURCE_GEOGCRS_CODE] = " & Str(NewCRS.SourceGeographicCRS.Code)
                    expression = "[COORD_REF_SYS_CODE] = " & Str(NewCRS.SourceGeographicCRS.Code)
                    Dim coordRefSysParameters = ds.Tables("CoordRefSys").Select(expression)
                    'COORD_REF_SYS_CODE COORD_REF_SYS_NAME AREA_OF_USE_CODE COORD_REF_SYS_KIND COORD_SYS_CODE DATUM_CODE SOURCE_GEOGCRS_CODE PROJECTION_CONV_CODE CRS_SCOPE REMARKS DEPRECATED
                    If coordRefSysParameters.Count > 0 Then
                        NewCRS.SourceGeographicCRS.Name = coordRefSysParameters(0).Item("COORD_REF_SYS_NAME")
                        NewCRS.SourceGeographicCRS.Author = "EPSG"
                        NewCRS.SourceGeographicCRS.Type = coordRefSysParameters(0).Item("COORD_REF_SYS_KIND")
                    End If

                    'Add Projection Operation information --------------------------------------------------------------------------------
                    If IsDBNull(ds.Tables("CoordRefSys").Rows(RowNo).Item("PROJECTION_CONV_CODE")) Then
                        NewCRS.Projection.Code = 0
                    Else
                        NewCRS.Projection.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("PROJECTION_CONV_CODE")
                    End If

                    expression = "[COORD_OP_CODE] = " & Str(NewCRS.Projection.Code)
                    Dim coordOpParameters = ds.Tables("CoordOp").Select(expression)
                    'Table fields: COORD_OP_CODE, COORD_OP_NAME, COORD_OP_TYPE, SOURCE_CRS_CODE, TARGET_CRS_CODE, COORD_TFM_VERSION, COORD_OP_VARIANT, AREA_OF_USE_CODE, COORD_OP_SCOPE,
                    'COORD_OP_ACCURACY, COORD_OP_METHOD_CODE, UOM_CODE_SOURCE_COORD_DIFF, UOM_CODE_TARGET_COORD_DIFF, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE,
                    'REVISION_DATE, SHOW_OPERATION, DEPRECATED
                    If coordOpParameters.Count > 0 Then
                        NewCRS.Projection.Name = coordOpParameters(0).Item("COORD_OP_NAME")
                        NewCRS.Projection.Author = "EPSG"
                        NewCRS.Projection.Type = coordOpParameters(0).Item("COORD_OP_TYPE")
                        NewCRS.ProjectionMethod.Code = coordOpParameters(0).Item("COORD_OP_METHOD_CODE")
                    End If

                    'Add Projection Method information --------------------------------------------------------------------------------
                    expression = "[COORD_OP_METHOD_CODE] = " & Str(NewCRS.ProjectionMethod.Code)
                    Dim methodParameters = ds.Tables("CoordOpMethod").Select(expression)
                    'Table fields: COORD_OP_METHOD_CODE, COORD_OP_METHOD_NAME, REVERSE_OP, FORMULA, EXAMPLE, REMARKS, INFORMATION_SOURCE, DATA_SOURCE, REVISION_DATE, CHANGE_ID, DEPRECATED
                    If methodParameters.Count > 0 Then
                        NewCRS.ProjectionMethod.Name = methodParameters(0).Item("COORD_OP_METHOD_NAME")
                        NewCRS.ProjectionMethod.Author = "EPSG"
                    End If

                    'Add Area of Use information -----------------------------------------------------------------------------------------
                    NewCRS.Area.Code = ds.Tables("CoordRefSys").Rows(RowNo).Item("AREA_OF_USE_CODE")
                    expression = "[AREA_CODE] = " & Str(NewCRS.Area.Code)
                    Dim areaOfUseParameters = ds.Tables("Area").Select(expression)
                    If areaOfUseParameters.Count > 0 Then
                        NewCRS.Area.Name = areaOfUseParameters(0).Item("AREA_NAME").ToString
                        NewCRS.Area.Author = "EPSG"
                    End If
                    List.Add(NewCRS)
                End If
            Next
            RaiseEvent Message("Finished." & vbCrLf)
            myConnection.Close()
        End Sub

        'Function to return the list of Projected CRSs as an XDocument
        Public Function ToXDoc() As System.Xml.Linq.XDocument

            Dim XDoc = <?xml version="1.0" encoding="utf-8"?>
                       <!---->
                       <!--Projected Coordinate Reference System List File-->
                       <ProjectedCRSList>
                           <CreationDate><%= Format(CreationDate, "d-MMM-yyyy H:mm:ss") %></CreationDate>
                           <LastEditDate><%= Format(LastEditDate, "d-MMM-yyyy H:mm:ss") %></LastEditDate>
                           <Description><%= Description %></Description>
                           <!---->
                           <%= From item In List _
                                   Select _
                                 <ProjectedCRS>
                                     <Name><%= item.Name %></Name>
                                     <Author><%= item.Author %></Author>
                                     <Code><%= item.Code %></Code>
                                     <Selected><%= item.Selected %></Selected>
                                     <Comments><%= item.Comments %></Comments>
                                     <Scope><%= item.Scope %></Scope>
                                     <Deprecated><%= item.Deprecated %></Deprecated>
                                     <AliasNames>
                                         <%= From nameItem In item.AliasName _
                                             Select _
                                             <AliasName><%= nameItem %></AliasName>
                                         %>
                                     </AliasNames>
                                     <CoordinateSystem>
                                         <Name><%= item.CoordinateSystem.Name %></Name>
                                         <Author><%= item.CoordinateSystem.Author %></Author>
                                         <Code><%= item.CoordinateSystem.Code %></Code>
                                         <Type><%= item.CoordinateSystem.Type %></Type>
                                     </CoordinateSystem>
                                     <SourceGeographicCRS>
                                         <Name><%= item.SourceGeographicCRS.Name %></Name>
                                         <Author><%= item.SourceGeographicCRS.Author %></Author>
                                         <Code><%= item.SourceGeographicCRS.Code %></Code>
                                         <Type><%= item.SourceGeographicCRS.Type %></Type>
                                     </SourceGeographicCRS>
                                     <Projection>
                                         <Name><%= item.Projection.Name %></Name>
                                         <Author><%= item.Projection.Author %></Author>
                                         <Code><%= item.Projection.Code %></Code>
                                         <Type><%= item.Projection.Type %></Type>
                                     </Projection>
                                     <ProjectionMethod>
                                         <Name><%= item.ProjectionMethod.Name %></Name>
                                         <Author><%= item.ProjectionMethod.Author %></Author>
                                         <Code><%= item.ProjectionMethod.Code %></Code>
                                         <Type><%= item.ProjectionMethod.Type %></Type>
                                     </ProjectionMethod>
                                     <AreaOfUse>
                                         <Name><%= item.Area.Name %></Name>
                                         <Author><%= item.Area.Author %></Author>
                                         <Code><%= item.Area.Code %></Code>
                                     </AreaOfUse>
                                 </ProjectedCRS>
                           %>
                           <!---->
                       </ProjectedCRSList>
            Return XDoc
        End Function

        'Add a user to the list
        Public Sub AddUser()
            _nUsers += 1
            If _nUsers = 1 Then
                LoadFile()
            End If
        End Sub

        'Remove a user from the list.
        Public Sub RemoveUser()
            If _nUsers > 0 Then
                _nUsers -= 1
                If _nUsers = 0 Then 'The list can be cleared.
                    List.Clear()
                End If
            End If
        End Sub

#End Region 'Methods -----------------------------------------------------------------------------------------------------

#Region "Events" '--------------------------------------------------------------------------------------------------------

        Event ErrorMessage(ByVal Message As String) 'Send an error message.
        Event Message(ByVal Message As String) 'Send a normal message.

#End Region 'Events ------------------------------------------------------------------------------------------------------

    End Class 'ProjectedCRSList --------------------------------------------------------------------------------------------------------------------------------------------------------------


'End Class 'Coordinates

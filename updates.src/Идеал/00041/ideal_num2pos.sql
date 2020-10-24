if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Num2Pos]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[Num2Pos]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE FUNCTION Num2Pos
(
@val money, @pos money
)  
RETURNS money 
 AS  
BEGIN 
if (@val - @pos) = 0
return 1
else
return 0
return null
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


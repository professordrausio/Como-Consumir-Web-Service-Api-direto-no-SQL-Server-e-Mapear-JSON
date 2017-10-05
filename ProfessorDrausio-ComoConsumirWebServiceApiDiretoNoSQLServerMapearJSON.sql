-- ===================================================================================
-- Author.......: Drausio Henrique Chiarotti
-- Create date..: 03/10/2017
-- Description..: Como Consumir Web Service Api direto no SQL Server e Mapear JSON
-- Canal Youtube: https://www.youtube.com/user/professordrausio
-- Aula.........: https://www.youtube.com/watch?v=yhGcfYbNGP0
-- LinkeIn......: https://www.linkedin.com/in/drausiohenriquechiarotti/
-- ===================================================================================
/*
--Habilitar as Stored Procedures de Ole Automation
sp_configure 'show advanced options', 1;
GO
RECONFIGURE;
GO
sp_configure 'Ole Automation Procedures', 1;
GO
RECONFIGURE;
GO
*/

DECLARE 
	@intToken INT,
	
	@vchEndereco VARCHAR(MAX),
	@vchURL AS VARCHAR(MAX),
	@vchJSON AS VARCHAR(8000),
	@vchStatus AS VARCHAR(MAX),
	@intQtdeResultados AS INT,
	--Mapear JSON
	@vchJSONResults AS VARCHAR(MAX),
	@vchJSONResultsGeometry AS VARCHAR(MAX),
	@vchJSONResultsLocation AS VARCHAR(MAX),
	@fltLatitude FLOAT,
	@fltLongitude FLOAT;
	
	SET @vchEndereco = 'Av. Alan Kardec, 1451 - Centro, Bebedouro - SP';
	SET @vchURL = 'https://maps.googleapis.com/maps/api/geocode/json?address=' + @vchEndereco;

	--OLE Automation é mecanismo para a comunicação entre processos baseado em Component Object Model (COM)
	--Is the returned object token, and must be a local variable of data type int. This object token identifies the created OLE object and is used in calls to the other OLE Automation stored procedures.
	EXEC sp_OACreate 'MSXML2.XMLHTTP', @intToken OUT;
	--Chamada para método OPEN
	EXEC sp_OAMethod @intToken, 'open', NULL, 'get', @vchURL, 'false';
	--Chamada para método SEND
	EXEC sp_OAMethod @intToken, 'send';
	--Chamada para método RESPONSE TEXT
	EXEC sp_OAMethod @intToken, 'responseText', @vchJSON OUTPUT;

	--Site para visualizar JSON http://jsonviewer.stack.hu/
	SELECT @vchJSON;
	
	--Mapear o JSON.

	--Checar se é um JSON válido
	IF (ISJSON(@vchJSON) = 1)
	BEGIN
		--Verificar se o status é OK (sucesso), ou seja, se a chamada foi realizada com sucesso
		SET @vchStatus = (SELECT TOP 1 [value] FROM OPENJSON(@vchJSON) WHERE [key] = 'status');
			
		IF (@vchStatus = 'OK')
		BEGIN
			
			--Verificar a quantidade de resultados dentro do JSON.
			SET @intQtdeResultados = (SELECT COUNT([key]) FROM OPENJSON(@vchJSON, '$.results'));
				
			IF (@intQtdeResultados = 1)
			BEGIN
				
				SET @vchJSONResults = (SELECT TOP 1 [value] FROM OPENJSON(@vchJSON, '$.results'));
					
				SET @vchJSONResultsGeometry = (SELECT TOP 1 [value] FROM OPENJSON(@vchJSONResults) WHERE [key] = 'geometry');
					
				SET @vchJSONResultsLocation = (SELECT TOP 1 [value] FROM OPENJSON(@vchJSONResultsGeometry) WHERE [key] = 'location');
					
				SET @fltLatitude = (SELECT top 1 [value] FROM OPENJSON(@vchJSONResultsLocation) WHERE [key] = 'lat');
					
				SET @fltLongitude = (SELECT top 1 [value] FROM OPENJSON(@vchJSONResultsLocation) WHERE [key] = 'lng');

				SELECT
					@fltLatitude AS Latitude,
					@fltLongitude AS Longitude,
					'https://www.google.com/maps/search/?api=1&query=' + CAST(@fltLatitude AS VARCHAR) + ',' + CAST(@fltLongitude AS VARCHAR);
			END

		END

	END

	EXEC sp_OADestroy @intToken;	
	

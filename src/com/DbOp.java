package com;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.core.LoggerContext;


/*
 * Copyright 2004-2005 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/**
 * 
 * @author SITO3
 *
 */
public class DbOp 
{
	private Logger logger = null;
	private Connection dbConn = null;
	private String jdbcURL = new String();
	private String jdbcDriver = new String();
	/**
	 * Database object,initialize db connection
	 * @param logger Message logger
	 * @throws Exception
	 */
	public DbOp(Logger logger) throws Exception 
	{
		jdbcDriver = "com.mysql.jdbc.Driver";
		jdbcURL = "jdbc:mysql://bmetis2/operator_roster?" +
                                   "user=cstsang&password=30Dec2010";
		Class.forName(jdbcDriver);
		dbConn = DriverManager.getConnection(jdbcURL);
		this.logger=logger;
	}
	
	/*public int upDateUserInfo(User user) 
	{
		int result=-1;
		ResultSet rs = null;
		PreparedStatement stmt = null;
		String sql="update user set password=?,";
		sql+="quota=?,diskSpaceUsed=?,";
		sql+="userLocale=?,downloadSpeedLimit=?,";
		sql+="uploadSpeedLimit=? where user_name=?";
		
		try 
		{
			stmt=dbConn.prepareStatement(sql);
			stmt.setString(1, user.getPassword());
			stmt.setDouble(2, user.getQuota());
			stmt.setDouble(3, user.getDiskSpaceUsed());
			stmt.setString(4,user.getUserLocale());
			stmt.setDouble(5, user.getDownloadSpeedLitmit());
			stmt.setDouble(6, user.getUploadSpeedLitmit());
			stmt.setString(7, user.getName());
			result=stmt.executeUpdate();
		} 
		catch (SQLException e) 
		{
			e.printStackTrace();
		}
		finally
		{
			releaseResource(rs, stmt);
		}
		return result;
	}*/
	/**
	 * Close db connection
	 * @throws Exception
	 */
	public void close() throws Exception 
	{
		dbConn.close();
		dbConn = null;
	}
	/**
	 * Release resource for 
	 * @param r ResultSet object
	 * @param s PreparedStatement object
	 */
	private void releaseResource(ResultSet r, PreparedStatement s) 
	{
		if (r != null) 
		{
			try 
			{
				r.close();
			} 
			catch (SQLException e) 
			{
				e.printStackTrace();
			}
		}
		if (s != null) 
		{
			try 
			{
				s.close();
			} 
			catch (SQLException e) 
			{
				e.printStackTrace();
			}
		}
		r = null;
		s = null;
	}
	public static void main(String[] args) 
	{
		Logger logger = LogManager.getLogger(DbOp.class.getName());
		logger.debug("Log4j2 is ready.");
		try 
		{
			DbOp dbo=new DbOp(logger);
			dbo.close();
		} 
		catch (Exception e) 
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}

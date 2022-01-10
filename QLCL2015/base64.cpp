#include "stdafx.h"
#include "base64.h"
#include <iostream>

static const std::string base64_chars = 
			 "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
			 "abcdefghijklmnopqrstuvwxyz"
			 "0123456789+/";


static inline bool is_base64(unsigned char c) {
  return (isalnum(c) || (c == '+') || (c == '/'));
}

std::string base64_encode(unsigned char const* bytes_to_encode, unsigned int in_len) {
  std::string ret;
  int i = 0;
  int j = 0;
  unsigned char char_array_3[3];
  unsigned char char_array_4[4];

  while (in_len--) {
	char_array_3[i++] = *(bytes_to_encode++);
	if (i == 3) {
	  char_array_4[0] = (char_array_3[0] & 0xfc) >> 2;
	  char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
	  char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
	  char_array_4[3] = char_array_3[2] & 0x3f;

	  for(i = 0; (i <4) ; i++)
		ret += base64_chars[char_array_4[i]];
	  i = 0;
	}
  }

  if (i)
  {
	for(j = i; j < 3; j++)
	  char_array_3[j] = '\0';

	char_array_4[0] = ( char_array_3[0] & 0xfc) >> 2;
	char_array_4[1] = ((char_array_3[0] & 0x03) << 4) + ((char_array_3[1] & 0xf0) >> 4);
	char_array_4[2] = ((char_array_3[1] & 0x0f) << 2) + ((char_array_3[2] & 0xc0) >> 6);
	char_array_4[3] =   char_array_3[2] & 0x3f;

	for (j = 0; (j < i + 1); j++)
	  ret += base64_chars[char_array_4[j]];

	while((i++ < 3))
	  ret += '=';

  }

  return ret;

}


std::string base64_decode(std::string const& encoded_string) {
  int in_len = encoded_string.size();
  int i = 0;
  int j = 0;
  int in_ = 0;
  unsigned char char_array_4[4], char_array_3[3];
  std::string ret;

  while (in_len-- && ( encoded_string[in_] != '=') && is_base64(encoded_string[in_])) {
	char_array_4[i++] = encoded_string[in_]; in_++;
	if (i ==4) {
	  for (i = 0; i <4; i++)
		char_array_4[i] = base64_chars.find(char_array_4[i]);

	  char_array_3[0] = ( char_array_4[0] << 2       ) + ((char_array_4[1] & 0x30) >> 4);
	  char_array_3[1] = ((char_array_4[1] & 0xf) << 4) + ((char_array_4[2] & 0x3c) >> 2);
	  char_array_3[2] = ((char_array_4[2] & 0x3) << 6) +   char_array_4[3];

	  for (i = 0; (i < 3); i++)
		ret += char_array_3[i];
	  i = 0;
	}
  }

  if (i) {
	for (j = i; j <4; j++)
	  char_array_4[j] = 0;

	for (j = 0; j <4; j++)
	  char_array_4[j] = base64_chars.find(char_array_4[j]);

	char_array_3[0] = (char_array_4[0] << 2) + ((char_array_4[1] & 0x30) >> 4);
	char_array_3[1] = ((char_array_4[1] & 0xf) << 4) + ((char_array_4[2] & 0x3c) >> 2);
	char_array_3[2] = ((char_array_4[2] & 0x3) << 6) + char_array_4[3];

	for (j = 0; (j < i - 1); j++) ret += char_array_3[j];
  }

  return ret;
}


CString f_B64EncodeEX(CString schuoi)
{
	try
	{
		int nlen = schuoi.GetLength();

		CString fx[15];
		CString szval=L"",sGan=L"";		
		for (int i = 0; i < 3; i++) fx[i] = schuoi.Mid(i,1);
		szval.Format(L"%s%s%s%s",fx[2],fx[0],fx[1],schuoi.Right(nlen-3));

		for (int i = 2; i >= 0; i--) fx[2-i] = szval.Mid(nlen-i-1,1);
		sGan.Format(L"%s%s%s%s",
			szval.Left(nlen-3),fx[2],fx[0],fx[1]);

		CT2CA pszConvert (sGan);
		const std::string sss = pszConvert;
		std::string encoded = base64_encode(
			reinterpret_cast<const unsigned char*>(sss.c_str()), sss.length());
		CString szbase(encoded.c_str());

		nlen = szbase.GetLength();
		for (int i = 0; i < 6; i++) fx[i] = szbase.Mid(i,1);
		szval.Format(L"%s%s%s%s%s%s%s",
			fx[1],fx[3],fx[5],fx[2],fx[0],fx[4],szbase.Right(nlen-6));

		for (int i = 5; i >= 0; i--) fx[5-i] = szval.Mid(nlen-i-1,1);
		sGan.Format(L"%s%s%s%s%s%s%s",
			szval.Left(nlen-6),fx[3],fx[2],fx[0],fx[5],fx[1],fx[4]);

		CT2CA psz2 (sGan);
		const std::string ss2 = psz2;
		std::string enc2 = base64_encode(
			reinterpret_cast<const unsigned char*>(ss2.c_str()), ss2.length());
		CString szb2(enc2.c_str());

		nlen = szb2.GetLength();
		if(nlen<16)
		{
			for (int i = 0; i < 8; i++) fx[i] = szb2.Mid(i,1);
			szval.Format(L"%s%s%s%s%s%s%s%s%s",
				fx[2],fx[0],fx[7],fx[1],fx[6],fx[4],fx[3],fx[5],szb2.Right(nlen-8));

			for (int i = 7; i >= 0; i--) fx[7-i] = szval.Mid(nlen-i-1,1);
			sGan.Format(L"%s%s%s%s%s%s%s%s%s",
				szval.Left(nlen-8),fx[1],fx[3],fx[6],fx[4],fx[7],fx[0],fx[5],fx[2]);
		}
		else
		{
			for (int i = 0; i < 12; i++) fx[i] = szb2.Mid(i,1);
			szval.Format(L"%s%s%s%s%s%s%s%s%s%s%s%s%s",
				fx[9],fx[7],fx[0],fx[4],fx[11],fx[1],fx[2],fx[10],fx[3],fx[5],fx[6],fx[8],szb2.Right(nlen-12));

			for (int i = 11; i >= 0; i--) fx[11-i] = szval.Mid(nlen-i-1,1);
			sGan.Format(L"%s%s%s%s%s%s%s%s%s%s%s%s%s",
				szval.Left(nlen-12),fx[8],fx[9],fx[6],fx[10],fx[5],fx[2],fx[0],fx[3],fx[11],fx[1],fx[7],fx[4]);
		}		

		return sGan;
	}
	catch(...){return schuoi;}

	return schuoi;
}


CString f_B64DecodeEX(CString sChuoi)
{
	try
	{
		int nlen = sChuoi.GetLength();

		CString fx[15];
		CString szval=L"",sGan=L"";		
		if(nlen<16)
		{
			for (int i = 7; i >= 0; i--) fx[7-i] = sChuoi.Mid(nlen-i-1,1);
				szval.Format(L"%s%s%s%s%s%s%s%s%s",
					sChuoi.Left(nlen-8),fx[5],fx[0],fx[7],fx[1],fx[3],fx[6],fx[2],fx[4]);

				for (int i = 0; i < 8; i++) fx[i] = szval.Mid(i,1);
				sGan.Format(L"%s%s%s%s%s%s%s%s%s",
					fx[1],fx[3],fx[0],fx[6],fx[5],fx[7],fx[4],fx[2],szval.Right(nlen-8));
		}
		else
		{
			for (int i = 11; i >= 0; i--) fx[11-i] = sChuoi.Mid(nlen-i-1,1);
				szval.Format(L"%s%s%s%s%s%s%s%s%s%s%s%s%s",
					sChuoi.Left(nlen-12),fx[6],fx[9],fx[5],fx[7],fx[11],fx[4],fx[2],fx[10],fx[0],fx[1],fx[3],fx[8]);

			for (int i = 0; i < 12; i++) fx[i] = szval.Mid(i,1);
				sGan.Format(L"%s%s%s%s%s%s%s%s%s%s%s%s%s",
					fx[2],fx[5],fx[6],fx[8],fx[3],fx[9],fx[10],fx[1],fx[11],fx[0],fx[7],fx[4],szval.Right(nlen-12));
		}			

		CT2CA pszConvert (sGan);
		const std::string sss = pszConvert;
		std::string decoded = base64_decode(sss);
		CString szbase(decoded.c_str());

		nlen = szbase.GetLength();
		for (int i = 5; i >= 0; i--) fx[5-i] = szbase.Mid(nlen-i-1,1);
		szval.Format(L"%s%s%s%s%s%s%s",
			szbase.Left(nlen-6),fx[2],fx[4],fx[1],fx[0],fx[5],fx[3]);

		for (int i = 0; i < 6; i++) fx[i] = szval.Mid(i,1);
		sGan.Format(L"%s%s%s%s%s%s%s",
			fx[4],fx[0],fx[3],fx[1],fx[5],fx[2],szval.Right(nlen-6));

		CT2CA psz2 (sGan);
		const std::string ss2 = psz2;
		std::string deco2 = base64_decode(ss2);
		CString szb2(deco2.c_str());

		nlen = szb2.GetLength();
		for (int i = 2; i >= 0; i--) fx[2-i] = szb2.Mid(nlen-i-1,1);
		szval.Format(L"%s%s%s%s",szb2.Left(nlen-3),fx[1],fx[2],fx[0]);

		for (int i = 0; i < 3; i++) fx[i] = szval.Mid(i,1);
			sGan.Format(L"%s%s%s%s",fx[1],fx[2],fx[0],szval.Right(nlen-3));

		return sGan;
	}
	catch(...){return sChuoi;}

	return sChuoi;
}
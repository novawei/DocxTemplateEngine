//
//  DocxTemplateEngine.h
//  DocxTemplateEngine
//
//  Created by WeiXinxing on 16/8/12.
//  Copyright © 2016年 nfs. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface DocxTemplateEngine : NSObject

@property (nonatomic, assign) NSSearchPathDirectory workspaceDirectory; // NSDocumentDirectory
@property (nonatomic, strong) NSString *workspaceName; // DocxWorkspace

+ (instancetype)sharedEngine;

/**
 *  create docx (word document)
 *
 *  @param name    filename, e.g. name=@"test", the docx's filename will be "test.docx"
 *  @param data    dictionary with key and value
 *  @param tplname template name (in bundle)
 *
 *  @return docx path if success, otherwise nil
 */
- (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname;
+ (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname;

@end

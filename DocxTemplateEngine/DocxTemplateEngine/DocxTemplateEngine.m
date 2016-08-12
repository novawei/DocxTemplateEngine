//
//  DocxTemplateEngine.m
//  DocxTemplateEngine
//
//  Created by WeiXinxing on 16/8/12.
//  Copyright © 2016年 nfs. All rights reserved.
//

#import "DocxTemplateEngine.h"
#import "SSZipArchive.h"

@interface DocxTemplateEngine ()

@property (nonatomic, strong) NSString *workspaceDirectoryPath;
- (void)createWorkspaceIfNeeded;

@end

@implementation DocxTemplateEngine

+ (instancetype)sharedEngine
{
    static DocxTemplateEngine *instance;
    static dispatch_once_t onceToken;
    dispatch_once(&onceToken, ^{
        instance = [[DocxTemplateEngine alloc] init];
    });
    return instance;
}

- (instancetype)init
{
    if (self = [super init]) {
        self.workspaceDirectory = NSDocumentDirectory;
        self.workspaceName = @"DocxWorkspace";
        self.placeholderFormat = @"{{%@}}";
        
        [self createWorkspaceIfNeeded];
    }
    return self;
}

- (void)createWorkspaceIfNeeded
{
    self.workspaceDirectoryPath = NSSearchPathForDirectoriesInDomains(self.workspaceDirectory, NSUserDomainMask, YES)[0];
    self.workspaceDirectoryPath = [self.workspaceDirectoryPath stringByAppendingPathComponent:self.workspaceName];
    if ([[NSFileManager defaultManager] fileExistsAtPath:self.workspaceDirectoryPath]) {
        return;
    }
    
    [[NSFileManager defaultManager] createDirectoryAtPath:self.workspaceDirectoryPath
                              withIntermediateDirectories:YES
                                               attributes:nil
                                                    error:nil];
}

- (NSString *)createDocxWithName:(NSString *)name 
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname
{
    return [self createDocxWithName:name
                               data:data
                      usingTemplate:tplname
                  andTemplateConfig:tplname];
}

+ (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname
{
    return [[DocxTemplateEngine sharedEngine] createDocxWithName:name
                                                            data:data
                                                   usingTemplate:tplname];
}

- (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname
               andTemplateConfig:(NSString *)confname
{
    NSString *tplBundlePath = [[NSBundle mainBundle] pathForResource:tplname ofType:@"docx"];
    NSString *tplWorkspacePath = [self.workspaceDirectoryPath stringByAppendingFormat:@"/%@.docx", tplname];
    if ([[NSFileManager defaultManager] fileExistsAtPath:tplWorkspacePath]) {
        [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
    }
    [[NSFileManager defaultManager] copyItemAtPath:tplBundlePath toPath:tplWorkspacePath error:nil];
    // create tmp directory
    NSString *tmpDirectoryPath = [self.workspaceDirectoryPath stringByAppendingPathComponent:@"tmp"];
    if ([[NSFileManager defaultManager] fileExistsAtPath:tmpDirectoryPath]) {
        [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
    }
    [[NSFileManager defaultManager] createDirectoryAtPath:tmpDirectoryPath
                              withIntermediateDirectories:YES
                                               attributes:nil
                                                    error:nil];

    BOOL docxRet = [SSZipArchive unzipFileAtPath:tplWorkspacePath toDestination:tmpDirectoryPath];
    if (!docxRet) { // cleanup
        [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
        [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
        return nil;
    }
    
    docxRet = NO;
    // word/document.xml
    NSString *contentPath = [tmpDirectoryPath stringByAppendingPathComponent:@"word/document.xml"];
    NSMutableString *content = [[NSMutableString alloc] initWithContentsOfFile:contentPath encoding:NSUTF8StringEncoding error:nil];
    // Replace placeholders
    NSString *confPath = [[NSBundle mainBundle] pathForResource:confname ofType:@"json"];
    if (content.length > 0 && [[NSFileManager defaultManager] fileExistsAtPath:confPath]) {
        BOOL contentChanged = NO;
        
        NSData *jsonData = [[NSData alloc] initWithContentsOfFile:confPath];
        NSArray *placeholders = [NSJSONSerialization JSONObjectWithData:jsonData options:0 error:nil];
        
        for (NSString *key in placeholders) {
            NSString *value = data[key];
            if (value) {
                contentChanged = YES;
                NSString *placeholder = [[NSString alloc] initWithFormat:self.placeholderFormat, key];
                [content replaceOccurrencesOfString:placeholder
                                         withString:value
                                            options:0
                                              range:NSMakeRange(0, content.length)];
            }
        }
        
        if (contentChanged) {
            docxRet = [content writeToFile:contentPath atomically:YES encoding:NSUTF8StringEncoding error:nil];
        }
    }
    
    if (!docxRet) { // cleanup
        [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
        [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
        return nil;
    }
    
    // zip
    NSString *zipPath = [self.workspaceDirectoryPath stringByAppendingFormat:@"/%@.zip", name];
    docxRet = [SSZipArchive createZipFileAtPath:zipPath withContentsOfDirectory:tmpDirectoryPath];
    if (!docxRet) { // cleanup
        [[NSFileManager defaultManager] removeItemAtPath:zipPath error:nil];
        [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
        [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
        return nil;
    }
    
    // rename
    NSString *docxPath = [self.workspaceDirectoryPath stringByAppendingFormat:@"/%@.docx", name];
    if ([[NSFileManager defaultManager] fileExistsAtPath:docxPath]) {
        [[NSFileManager defaultManager] removeItemAtPath:docxPath error:nil];
    }
    [[NSFileManager defaultManager] moveItemAtPath:zipPath toPath:docxPath error:nil];
    // cleanup
    [[NSFileManager defaultManager] removeItemAtPath:tplWorkspacePath error:nil];
    [[NSFileManager defaultManager] removeItemAtPath:tmpDirectoryPath error:nil];
    
    return docxPath;
}

+ (NSString *)createDocxWithName:(NSString *)name
                            data:(NSDictionary *)data
                   usingTemplate:(NSString *)tplname
               andTemplateConfig:(NSString *)confname
{
    return [[DocxTemplateEngine sharedEngine] createDocxWithName:name
                                                            data:data
                                                   usingTemplate:tplname
                                               andTemplateConfig:confname];
}

@end
